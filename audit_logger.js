/******************************
 * GOOGLE SHEETS FULL AUDIT LOGGER (DIFF AWARE)
 * Tracks:
 * - Cell edits (values + formulas)
 * - Bulk edits (paste, drag, clear ranges)
 * - Structural changes (rows, columns, sheets, renames)
 * Logs with exact row/column/sheet info
 ******************************/

function onOpen() {
  saveSheetStructure();
}

function onEdit(e) {
  try {
    logEditEvent(e);
    saveSheetStructure();
  } catch (err) {
    console.error("onEdit error: " + err);
  }
}

function onChange(e) {
  try {
    logChangeEvent(e);
    saveSheetStructure();
  } catch (err) {
    console.error("onChange error: " + err);
  }
}

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
      message = `${user} renamed a sheet at ${timestamp}`;
      break;

    default:
      message = `${user} performed a '${changeType}' action at ${timestamp}`;
  }

  if (message) {
    logSheet.appendRow([timestamp, user, changeType, message]);
  }
}

/** Detect exact row changes */
function detectRowChange(oldStruct, newStruct, user, timestamp, changeType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
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

/** Detect exact column changes */
function detectColumnChange(oldStruct, newStruct, user, timestamp, changeType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
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

/** Capture sheet structure */
function captureCurrentStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var logSpreadsheet = SpreadsheetApp.openById("1BanMsUh4wqk_URQE98sAvl0VGvFTZK8P4sluDrNC-90"); 
  var logSheet = logSpreadsheet.getSheetByName("Logs");

  // If "Logs" sheet doesn’t exist yet, create it
  if (!logSheet) {
    logSheet = logSpreadsheet.insertSheet("Logs");
    logSheet.appendRow(["Timestamp", "User", "Action Type", "Details"]);
  }

  return logSheet;
}


function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}
