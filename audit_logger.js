/******************************
 * GOOGLE SHEETS FULL AUDIT LOGGER (INDUSTRY READY)
 * Standalone version (secure)
 * Tracks:
 * - Cell edits (values + formulas)
 * - Bulk edits (paste, drag, clear ranges)
 * - Structural changes (rows, columns, sheets, renames)
 * Logs with exact row/column/sheet info
 ******************************/

// IDs (replace these with your own)
var MAIN_SHEET_ID = "PUT-YOUR-MAIN-SHEET-ID-HERE";
var LOG_SHEET_ID  = "PUT-YOUR-PRIVATE-LOG-SHEET-ID-HERE";

/******************************
 * TRIGGER HANDLERS
 ******************************/

function onMainEdit(e) {
  try {
    logEditEvent(e);
    saveSheetStructure();
  } catch (err) {
    console.error("onMainEdit error: " + err);
  }
}

function onMainChange(e) {
  try {
    logChangeEvent(e);
    saveSheetStructure();
  } catch (err) {
    console.error("onMainChange error: " + err);
  }
}

/******************************
 * LOGGING FUNCTIONS
 ******************************/

/** Handle cell edits (values, formulas, ranges) */
function logEditEvent(e) {
  var logSheet = getLogSheet();
  var user = Session.getActiveUser().getEmail() || "Unknown User";
  var sheetName = e.range.getSheet().getName();
  var timestamp = formatDate(new Date());

  var range = e.range;
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  // Bulk edit
  if (numRows > 1 || numCols > 1) {
    var firstCell = range.getCell(1, 1);
    var firstFormula = firstCell.getFormula();
    var firstValue = firstFormula ? "Formula: " + firstFormula : firstCell.getValue();
    if (firstValue === "") firstValue = "(cleared)";

    var message = `${user} updated range ${range.getA1Notation()} on '${sheetName}' with ${
      firstFormula ? "formulas" : "values"
    } (first cell: '${firstValue}') at ${timestamp}`;

    logSheet.appendRow([timestamp, user, "BULK_EDIT", message]);
    return;
  }

  // Single cell edit
  var cell = range.getA1Notation();
  var oldValue = e.oldValue !== undefined ? e.oldValue : "(blank)";
  var newValue = range.getFormula() ? "Formula: " + range.getFormula() : range.getValue();
  if (newValue === "") newValue = "(cleared)";

  var message = `${user} edited ${cell} on '${sheetName}' from '${oldValue}' to '${newValue}' at ${timestamp}`;
  logSheet.appendRow([timestamp, user, "EDIT", message]);
}

/** Handle structural changes */
function logChangeEvent(e) {
  var logSheet = getLogSheet();
  var user = Session.getActiveUser().getEmail() || "Unknown User";
  var timestamp = formatDate(new Date());
  var changeType = e.changeType || "UNKNOWN_CHANGE";

  var oldStruct = getOldStructure();
  var newStruct = captureCurrentStructure();

  var message;

  switch (changeType) {
    case "INSERT_ROW":
    case "REMOVE_ROW":
      message = detectRowChange(oldStruct, newStruct, user, timestamp, changeType);
      break;

    case "INSERT_COLUMN":
    case "REMOVE_COLUMN":
      message = detectColumnChange(oldStruct, newStruct, user, timestamp, changeType);
      break;

    case "INSERT_GRID":
      var addedSheet = Object.keys(newStruct).filter((s) => !oldStruct[s])[0];
      message = `${user} added a new sheet '${addedSheet}' at ${timestamp}`;
      break;

    case "REMOVE_GRID":
      var removedSheet = Object.keys(oldStruct).filter((s) => !newStruct[s])[0];
      message = `${user} deleted sheet '${removedSheet}' at ${timestamp}`;
      break;

    case "RENAME_SHEET":
      // Compare names
      var oldNames = Object.keys(oldStruct);
      var newNames = Object.keys(newStruct);
      var renamedFrom = oldNames.find(n => !newNames.includes(n));
      var renamedTo   = newNames.find(n => !oldNames.includes(n));
      if (renamedFrom && renamedTo) {
        message = `${user} renamed sheet '${renamedFrom}' to '${renamedTo}' at ${timestamp}`;
      } else {
        message = `${user} renamed a sheet at ${timestamp}`;
      }
      break;

    default:
      message = `${user} performed a '${changeType}' action at ${timestamp}`;
  }

  if (message) {
    logSheet.appendRow([timestamp, user, changeType, message]);
  }
}

/******************************
 * HELPERS
 ******************************/

function detectRowChange(oldStruct, newStruct, user, timestamp, changeType) {
  var sheet = SpreadsheetApp.openById(MAIN_SHEET_ID).getActiveSheet();
  var name = sheet.getName();
  var oldRows = oldStruct[name] ? oldStruct[name].rows : 0;
  var newRows = newStruct[name] ? newStruct[name].rows : 0;

  if (newRows > oldRows) {
    return `${user} inserted row(s) (approx. at index ${sheet.getActiveRange().getRow()}) in '${name}' at ${timestamp}`;
  } else if (newRows < oldRows) {
    return `${user} deleted row(s) (approx. at index ${sheet.getActiveRange().getRow()}) in '${name}' at ${timestamp}`;
  }
  return null;
}

function detectColumnChange(oldStruct, newStruct, user, timestamp, changeType) {
  var sheet = SpreadsheetApp.openById(MAIN_SHEET_ID).getActiveSheet();
  var name = sheet.getName();
  var oldCols = oldStruct[name] ? oldStruct[name].cols : 0;
  var newCols = newStruct[name] ? newStruct[name].cols : 0;

  if (newCols > oldCols) {
    return `${user} inserted column(s) (approx. at index ${sheet.getActiveRange().getColumn()}) in '${name}' at ${timestamp}`;
  } else if (newCols < oldCols) {
    return `${user} deleted column(s) (approx. at index ${sheet.getActiveRange().getColumn()}) in '${name}' at ${timestamp}`;
  }
  return null;
}

function captureCurrentStructure() {
  var ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
  var sheets = ss.getSheets();
  var structure = {};
  sheets.forEach(function (sh) {
    structure[sh.getName()] = {
      rows: sh.getMaxRows(),
      cols: sh.getMaxColumns()
    };
  });
  return structure;
}

function saveSheetStructure() {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty("sheetStructure", JSON.stringify(captureCurrentStructure()));
}

function getOldStructure() {
  var props = PropertiesService.getDocumentProperties();
  var data = props.getProperty("sheetStructure");
  return data ? JSON.parse(data) : {};
}

function getLogSheet() {
  var logSpreadsheet = SpreadsheetApp.openById(LOG_SHEET_ID);
  var logSheet = logSpreadsheet.getSheetByName("Logs");

  if (!logSheet) {
    logSheet = logSpreadsheet.insertSheet("Logs");
    logSheet.appendRow(["Timestamp", "User", "Action Type", "Details"]);
  }
  return logSheet;
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

/******************************
 * ONE-TIME SETUP
 ******************************/

// Run this manually once to create the triggers
function createTriggers() {
  var ss = SpreadsheetApp.openById(MAIN_SHEET_ID);

  // Watch edits
  ScriptApp.newTrigger("onMainEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Watch structural changes
  ScriptApp.newTrigger("onMainChange")
    .forSpreadsheet(ss)
    .onChange()
    .create();
}
