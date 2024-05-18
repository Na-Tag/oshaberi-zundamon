function setLog(contents) {
  const sheet = SpreadsheetApp.openById(SPREAD_SHEET_ID).getSheetByName("ログ");
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(lastRow + 1,1,1,3);
  range.setValues([contents]);
}
