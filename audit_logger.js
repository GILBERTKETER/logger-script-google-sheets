/******************************
 * GOOGLE SHEETS MULTI-SPREADSHEET AUDIT LOGGER
 * Tracks edits + structural changes from multiple sheets
 * Logs into one private spreadsheet, separate tabs per source
 ******************************/

// Central private log spreadsheet
var LOG_SHEET_ID = "PUT-YOUR-PRIVATE-LOG-SHEET-ID-HERE";

// Sheets to monitor
var MONITORED_SHEETS = [
  { id: "PUT-SHEET-ID-1-HERE", logName: "SheetA_Logs" },
  { id: "PUT-SHEET-ID-2-HERE", logName: "SheetB_Logs" },
  { id: "PUT-SHEET-ID-3-HERE", logName: "SheetC_Logs" }
];

/******************************
 * TRIGGER HANDLERS
 ******************************/

function onAnyEdit(e) {
  try {
    logEditEvent(e);
    saveSheetStructure(e.source.getId());
  } catch (err) {
    console.error("onAnyEdit error: " + err);
  }
}

function onAnyChange(e) {
  try {
    logChangeEvent(e);
    saveSheetStructure(e.source.getId());
  } catch (err) {
    console.error("onAnyChange error: " + err);
  }
}

/******************************
 * LOGGING FUNCTIONS
 ******************************/

function logEditEvent(e) {
  var sourceId = e.source.getId();
  var logSheet = getLogSheetForSource(sourceId);
  if (!logSheet) return; // Not a monitored sheet

  var user = Session.getActiveUser().getEmail() || "Unknown User";
  var sheetName = e.range.getSheet().getName();
  var timestamp = formatDate(new Date());

  var range = e.range;
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

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

  var cell = range.getA1Notation();
  var oldValue = e.oldValue !== undefined ? e.oldValue : "(blank)";
  var newValue = range.getFormula() ? "Formula: " + range.getFormula() : range.getValue();
  if (newValue === "") newValue = "(cleared)";

  var message = `${user} edited ${cell} on '${sheetName}' from '${oldValue}' to '${newValue}' at ${timestamp}`;
  logSheet.appendRow([timestamp, user, "EDIT", message]);
}

function logChangeEvent(e) {
  var sourceId = e.source.getId();
  var logSheet = getLogSheetForSource(sourceId);
  if (!logSheet) return;

  var user = Session.getActiveUser().getEmail() || "Unknown User";
  var timestamp = formatDate(new Date());
  var changeType = e.changeType || "UNKNOWN_CHANGE";

  var oldStruct = getOldStructure(sourceId);
  var newStruct = captureCurrentStructure(sourceId);

  var message;

  switch (changeType) {
    case "INSERT_ROW":
    case "REMOVE_ROW":
      message = detectRowChange(oldStruct, newStruct, user, timestamp, changeType, sourceId);
      break;

    case "INSERT_COLUMN":
    case "REMOVE_COLUMN":
      message = detectColumnChange(oldStruct, newStruct, user, timestamp, changeType, sourceId);
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
 * STRUCTURE + HELPERS
 ******************************/

function detectRowChange(oldStruct, newStruct, user, timestamp, changeType, sourceId) {
  var sheet = SpreadsheetApp.openById(sourceId).getActiveSheet();
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

function detectColumnChange(oldStruct, newStruct, user, timestamp, changeType, sourceId) {
  var sheet = SpreadsheetApp.openById(sourceId).getActiveSheet();
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

function captureCurrentStructure(sourceId) {
  var ss = SpreadsheetApp.openById(sourceId);
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

function saveSheetStructure(sourceId) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty("sheetStructure_" + sourceId, JSON.stringify(captureCurrentStructure(sourceId)));
}

function getOldStructure(sourceId) {
  var props = PropertiesService.getDocumentProperties();
  var data = props.getProperty("sheetStructure_" + sourceId);
  return data ? JSON.parse(data) : {};
}

function getLogSheetForSource(sourceId) {
  var mapping = MONITORED_SHEETS.find(s => s.id === sourceId);
  if (!mapping) return null;

  var logSpreadsheet = SpreadsheetApp.openById(LOG_SHEET_ID);
  var logSheet = logSpreadsheet.getSheetByName(mapping.logName);

  if (!logSheet) {
    logSheet = logSpreadsheet.insertSheet(mapping.logName);
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

function createTriggers() {
  MONITORED_SHEETS.forEach(function (entry) {
    var ss = SpreadsheetApp.openById(entry.id);

    // Watch edits
    ScriptApp.newTrigger("onAnyEdit")
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    // Watch structural changes
    ScriptApp.newTrigger("onAnyChange")
      .forSpreadsheet(ss)
      .onChange()
      .create();
  });
}
