/** Shared owner notification engine. */
function isApprovalRequired_(status) {
  const s = String(status || '').trim().toUpperCase();
  if (!s) return false;
  if (KALMA_NOTIFY_CONFIG.FINAL_STATUSES.indexOf(s) >= 0) return false;
  return KALMA_NOTIFY_CONFIG.APPROVAL_STATUSES.indexOf(s) >= 0;
}

function buildNotificationPayload_(input) {
  const now = new Date();
  const data = input || {};
  const moduleName = cleanText_(data.module || data.moduleName || 'UNKNOWN_MODULE');
  const sourceSheet = cleanText_(data.sourceSheet || data.sheet || 'UNKNOWN_SOURCE');
  const recordId = cleanText_(data.recordId || data.id || data.submissionId || 'UNKNOWN_RECORD_ID');
  const status = cleanText_(data.status || '').toUpperCase();
  const submittedBy = cleanText_(data.submittedBy || data.pic || data.user || 'UNKNOWN');
  const amount = normalizeAmount_(data.amount);
  const vendor = cleanText_(data.vendor || data.payee || data.customer || '');
  const description = cleanText_(data.description || data.note || data.summary || '');
  const spreadsheetUrl = cleanText_(data.spreadsheetUrl || getActiveSpreadsheetUrlSafe_() || '');
  const approvalLink = cleanText_(data.approvalLink || data.openUrl || spreadsheetUrl || '');
  const timestamp = data.timestamp instanceof Date ? data.timestamp : now;
  const channel = cleanText_(data.channel || 'Telegram');

  const payload = {
    timestamp,
    module: moduleName,
    sourceSheet,
    recordId,
    status,
    submittedBy,
    amount,
    vendor,
    description,
    channel,
    spreadsheetUrl,
    approvalLink
  };
  payload.notificationKey = generateNotificationKey_(payload);
  payload.messageText = buildTelegramMessageText_(payload);
  return payload;
}

function generateNotificationKey_(payload) {
  const p = payload || {};
  return [
    p.module,
    p.sourceSheet,
    p.recordId,
    p.status,
    p.channel || 'Telegram'
  ].map(v => cleanKeyPart_(v)).join('|');
}

function notificationAlreadySent_(notificationKey, spreadsheetId) {
  const sheet = getNotificationLogSheet_(spreadsheetId, false);
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const keyCol = headers.indexOf('Notification Key') + 1;
  const resultCol = headers.indexOf('Notification Result') + 1;
  if (!keyCol || !resultCol) return false;
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return values.some(row => String(row[keyCol - 1]) === String(notificationKey) && /^SENT$/i.test(String(row[resultCol - 1])));
}

function writeNotificationLog_(payload, result, errorMessage, spreadsheetId) {
  const sheet = getNotificationLogSheet_(spreadsheetId, true);
  const p = payload || {};
  const row = [
    formatTimestamp_(p.timestamp || new Date()),
    p.module || '',
    p.sourceSheet || '',
    p.recordId || '',
    p.status || '',
    p.submittedBy || '',
    p.amount || '',
    p.vendor || '',
    p.description || '',
    p.channel || 'Telegram',
    result || 'UNKNOWN',
    p.notificationKey || generateNotificationKey_(p),
    p.spreadsheetUrl || '',
    errorMessage || ''
  ];
  sheet.appendRow(row);
}

function notifyOwnerIfNeeded_(input) {
  const payload = buildNotificationPayload_(input);
  if (!isApprovalRequired_(payload.status)) {
    return { ok: true, skipped: true, reason: 'STATUS_NOT_APPROVAL_REQUIRED', notificationKey: payload.notificationKey };
  }

  const spreadsheetId = input && input.spreadsheetId ? input.spreadsheetId : '';
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { ok: false, preservedTransaction: true, notificationResult: 'FAILED', error: 'NOTIFICATION_LOCK_TIMEOUT: ' + e.message, notificationKey: payload.notificationKey };
  }

  try {
    if (notificationAlreadySent_(payload.notificationKey, spreadsheetId)) {
      return { ok: true, skipped: true, reason: 'DUPLICATE_ALREADY_SENT', notificationKey: payload.notificationKey };
    }

    const telegramResult = sendTelegramNotification_(payload);
    writeNotificationLog_(payload, telegramResult.ok ? 'SENT' : 'FAILED', telegramResult.error || '', spreadsheetId);
    return {
      ok: telegramResult.ok,
      preservedTransaction: true,
      notificationResult: telegramResult.ok ? 'Telegram Sent' : 'Telegram Failed',
      channel: 'Telegram',
      notificationKey: payload.notificationKey,
      responseCode: telegramResult.responseCode || '',
      error: telegramResult.error || ''
    };
  } catch (err) {
    try { writeNotificationLog_(payload, 'FAILED', err && err.message ? err.message : String(err), spreadsheetId); } catch (logErr) { console.error(logErr); }
    return {
      ok: false,
      preservedTransaction: true,
      notificationResult: 'Telegram Failed',
      channel: 'Telegram',
      notificationKey: payload.notificationKey,
      error: err && err.message ? err.message : String(err)
    };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function getNotificationStatusForRecord_(moduleName, sourceSheet, recordId, status, channel, spreadsheetId) {
  const payload = buildNotificationPayload_({ module: moduleName, sourceSheet, recordId, status, channel: channel || 'Telegram' });
  const sheet = getNotificationLogSheet_(spreadsheetId, false);
  if (!sheet) return { status: 'No Notification Log', notificationKey: payload.notificationKey };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'Not Sent', notificationKey: payload.notificationKey };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const keyCol = headers.indexOf('Notification Key') + 1;
  const resultCol = headers.indexOf('Notification Result') + 1;
  const errCol = headers.indexOf('Error Message') + 1;
  if (!keyCol || !resultCol) return { status: 'Notification Log Invalid', notificationKey: payload.notificationKey };
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    if (String(row[keyCol - 1]) === payload.notificationKey) {
      const res = String(row[resultCol - 1] || '');
      return {
        status: /^SENT$/i.test(res) ? 'Telegram Sent' : 'Telegram Failed',
        notificationKey: payload.notificationKey,
        error: errCol ? String(row[errCol - 1] || '') : ''
      };
    }
  }
  return { status: 'Not Sent', notificationKey: payload.notificationKey };
}
