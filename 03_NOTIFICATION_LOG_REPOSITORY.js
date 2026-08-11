/** NOTIFICATION_LOG repository. */
function getNotificationLogSheet_(spreadsheetId, createIfMissing) {
  let ss = null;
  if (spreadsheetId) ss = SpreadsheetApp.openById(spreadsheetId);
  else ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('NO_ACTIVE_SPREADSHEET_FOR_NOTIFICATION_LOG');
  let sheet = ss.getSheetByName(KALMA_NOTIFY_CONFIG.NOTIFICATION_LOG_SHEET);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(KALMA_NOTIFY_CONFIG.NOTIFICATION_LOG_SHEET);
    sheet.getRange(1, 1, 1, KALMA_NOTIFY_CONFIG.LOG_HEADERS.length).setValues([KALMA_NOTIFY_CONFIG.LOG_HEADERS]);
    sheet.setFrozenRows(1);
  }
  if (sheet) ensureNotificationLogHeaders_(sheet);
  return sheet;
}

function ensureNotificationLogHeaders_(sheet) {
  const required = KALMA_NOTIFY_CONFIG.LOG_HEADERS;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, required.length).setValues([required]);
    sheet.setFrozenRows(1);
    return;
  }
  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), required.length)).getValues()[0].map(String);
  let changed = false;
  required.forEach((h, idx) => {
    if (current[idx] !== h) changed = true;
  });
  if (changed) sheet.getRange(1, 1, 1, required.length).setValues([required]);
}
