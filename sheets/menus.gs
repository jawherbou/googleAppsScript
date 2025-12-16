/**
 * Adds custom menus to the Google Sheets UI and sets a predefined theme.
 */

function onOpen() {
  // add menus
  const ui = SpreadsheetApp.getUi()
  ui.createMenu('Approval')
    .addItem('Ask Approval', 'sendApprovalRequest')
    .addToUi();

   ui.createMenu('Helper Functions')
    .addSubMenu(
      ui.createMenu('Copy Tools')
        .addItem('Copy First Sheet', 'copyFirstSheetToNewSheet')
        .addItem('Copy Selection', 'copySelectedRangeToNewSheet')
        .addItem('Copy From External Spreadsheet', 'copyFromExternalSpreadsheet')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Utilities')
        .addItem('Show Sheet Info', 'showSpreadsheetInfo')
    )
    .addSeparator()
    .addItem('Help', 'showHelpDialog')
    .addToUi();

  // set themes
  setTheme();
}

function setTheme() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const predefinedThemesList = spreadSheet.getPredefinedSpreadsheetThemes();
  const theme = predefinedThemesList[1];

  theme.setConcreteColor(SpreadsheetApp.ThemeColorType.TEXT, 100,100,10);
  theme.setFontFamily(SpreadsheetApp.FontFamily.ARIAL);
  spreadSheet.setSpreadsheetTheme(theme);
}
