/**
 * Copies the currently selected range
 * into a newly created sheet.
 */
function copySelectedRangeToNewSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const selectedRange = activeSheet.getActiveRange();

  const rangeValues = selectedRange.getValues();
  createSheetAndPasteData(spreadsheet, rangeValues);
}

/**
 * Copies the first sheet in the current spreadsheet,
 * modifies currency values, and pastes into a new sheet.
 */
function copyFirstSheetToNewSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheets()[0];

  const sheetData = getEntireSheetData(sourceSheet);

  createSheetAndPasteData(spreadsheet, sheetData);
}

/**
 * Copies the first sheet from another spreadsheet
 * and pastes it into a new sheet in the active spreadsheet.
 */
function copyFromExternalSpreadsheet() {
  const externalSpreadsheetId = 'EXTERNAL_ID'; // replace with actual ID
  const externalSpreadsheet = SpreadsheetApp.openById(externalSpreadsheetId);
  const sourceSheet = externalSpreadsheet.getSheets()[0];

  const sheetData = getEntireSheetData(sourceSheet);

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  createSheetAndPasteData(activeSpreadsheet, sheetData);
}


/**
 * Retrieves all values from a sheet based on its
 * last row and last column.
 *
 * @param {Sheet} sheet - Google Sheets Sheet object
 * @returns {Array[][]} 2D array of sheet values
 */
function getEntireSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0 || lastColumn === 0) {
    return [];
  }

  const fullRange = sheet.getRange(1, 1, lastRow, lastColumn);
  return fullRange.getValues();
}

/**
 * Creates a new sheet and pastes data starting at A1.
 *
 * @param {Spreadsheet} spreadsheet - Target spreadsheet
 * @param {Array[][]} data - 2D array of values
 */
function createSheetAndPasteData(spreadsheet, data) {
  if (!data.length) return;

  const newSheet = spreadsheet.insertSheet();
  const targetRange = newSheet.getRange(1, 1, data.length, data[0].length);
  targetRange.setValues(data);
}

function showSpreadsheetInfo() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const message =
    `Spreadsheet Name: ${spreadsheet.getName()}\n` +
    `Sheets Count: ${spreadsheet.getSheets().length}`;

  SpreadsheetApp.getUi().alert('Spreadsheet Info', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showHelpDialog() {
  const html = HtmlService
    .createHtmlOutput('<p><b>Custom Functions</b><br>Use the menu to copy sheets or ranges.</p>')
    .setWidth(300)
    .setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(html, 'Help');
}

