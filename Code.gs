/**
 * @OnlyCurrentDoc Limits the script to only accessing the current spreadsheet.
 */

/**
 * Creates menu entries in the Sheets UI when the document is opened.
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Insert Cells')
    .addItem('Shift Down', 'shiftDown')
    .addItem('Shift Right', 'shiftRight')
    .addToUi();
  ui.createMenu('Delete Cells')
    .addItem('Shift Up', 'shiftUp')
    .addItem('Shift Left', 'shiftLeft')
    .addToUi();
}

/**
 * Insert blank cells and shift existing content down
 */
function shiftDown() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getActiveCell();
  var row = cell.getRow();
  var col = cell.getColumn();
  var range = sheet.getActiveRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var lastRow = sheet.getLastRow();
  try {
    sheet.getRange(row, col, lastRow - row + 1, numCols).moveTo(sheet.getRange(row + numRows, col, lastRow - row + 1, numCols));
    range.clearContent();
  }
  catch(err) {
    SpreadsheetApp.getUi().alert('Could not move data: Please make sure nobody else is editing the column(s) and try again.');
  }
}

/**
 * Insert blank cells and shift existing content to the right
 */
function shiftRight() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getActiveCell();
  var row = cell.getRow();
  var col = cell.getColumn();
  var range = sheet.getActiveRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var lastCol = sheet.getLastColumn();
  try {
    sheet.getRange(row, col, numRows, lastCol - col + 1).moveTo(sheet.getRange(row, col + numCols, numRows, lastCol - col + 1));
    range.clearContent();
  }
  catch (err) {
    SpreadsheetApp.getUi().alert('Could not move data: Please make sure nobody else is editing the row(s) and try again.');
  }
}

/**
 * Delete cells and shift up content from below
 */
function shiftUp() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getActiveCell();
  var row = cell.getRow();
  var col = cell.getColumn();
  var range = sheet.getActiveRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var lastRow = sheet.getLastRow();
  try {
    sheet.getRange(row + numRows, col, lastRow - row - numRows + 1, numCols).moveTo(sheet.getRange(row, col, lastRow - row - numRows + 1, numCols));
    sheet.getRange(lastRow - numRows + 1, col, numRows, numCols).clearContent();
  }
  catch(err) {
    if (row + numRows > lastRow) {
      SpreadsheetApp.getUi().alert('Select a different range: Cannot shift up from the last populated row of a sheet.');
    } else {
      SpreadsheetApp.getUi().alert('Could not move data: Please make sure nobody else is editing the column(s) and try again.');
    }
  }
}

/**
 * Delete cells and shift left content from the right
 */
function shiftLeft() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getActiveCell();
  var row = cell.getRow();
  var col = cell.getColumn();
  var range = sheet.getActiveRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var lastCol = sheet.getLastColumn();
  try {
    sheet.getRange(row, col + numCols, numRows, lastCol - col - numCols + 1).moveTo(sheet.getRange(row, col, numRows, lastCol - col - numCols + 1));
    sheet.getRange(row, lastCol - numCols + 1, numRows, numCols).clearContent();
  }
  catch(err) {
    if (col + numCols > lastCol) {
      SpreadsheetApp.getUi().alert('Select a different range: Cannot shift left from the last populated column of a sheet.');
    } else {
      SpreadsheetApp.getUi().alert('Could not move data: Please make sure nobody else is editing the row(s) and try again.');
    }
  }
}
