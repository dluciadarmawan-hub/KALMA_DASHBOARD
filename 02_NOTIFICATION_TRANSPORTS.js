/** Transport layer. Telegram only is active. */
function sendTelegramNotification_(payload) {
  if (!KALMA_NOTIFY_CONFIG.TELEGRAM.enabled) {
    return { ok: false, error: 'TELEGRAM_DISABLED' };
  }
  const token = PropertiesService.getScriptProperties().getProperty(KALMA_NOTIFY_CONFIG.TELEGRAM.tokenScriptProperty);
  if (!token) {
    return { ok: false, error: 'TELEGRAM_BOT_TOKEN_MISSING' };
  }
  const url = KALMA_NOTIFY_CONFIG.TELEGRAM.apiBaseUrl + encodeURIComponent(token) + '/sendMessage';
  const body = {
    chat_id: KALMA_NOTIFY_CONFIG.TELEGRAM.chatId,
    text: payload.messageText,
    parse_mode: KALMA_NOTIFY_CONFIG.TELEGRAM.parseMode,
    disable_web_page_preview: false
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  let parsed = null;
  try { parsed = JSON.parse(text); } catch (_) {}
  if (code >= 200 && code < 300 && (!parsed || parsed.ok !== false)) {
    return { ok: true, responseCode: code, raw: text };
  }
  return { ok: false, responseCode: code, error: 'TELEGRAM_HTTP_' + code + ': ' + text };
}

function sendEmailNotification_() {
  return { ok: false, skipped: true, disabled: true, channel: 'Email', error: 'EMAIL_DISABLED_BY_RELEASE_POLICY' };
}

function sendWhatsAppNotification_() {
  return { ok: false, skipped: true, disabled: true, channel: 'WhatsApp', error: 'WHATSAPP_DISABLED_BY_RELEASE_POLICY' };
}

function sendGoogleChatNotification_() {
  return { ok: false, skipped: true, disabled: true, channel: 'GoogleChat', error: 'GOOGLE_CHAT_DISABLED_BY_RELEASE_POLICY' };
}

function buildTelegramMessageText_(payload) {
  const p = payload || {};
  const lines = [
    '🔔 <b>APPROVAL REQUIRED</b>',
    '',
    '<b>Module:</b>', escapeHtml_(p.module || ''),
    '',
    '<b>Status:</b>', escapeHtml_(p.status || ''),
    '',
    '<b>Submitted By:</b>', escapeHtml_(p.submittedBy || ''),
    '',
    '<b>Vendor:</b>', escapeHtml_(p.vendor || '-'),
    '',
    '<b>Amount:</b>', formatRupiah_(p.amount),
    '',
    '<b>Description:</b>', escapeHtml_(p.description || '-'),
    '',
    '<b>Timestamp:</b>', escapeHtml_(formatTimestamp_(p.timestamp || new Date()))
  ];
  if (p.approvalLink) {
    lines.push('', '<b>Open:</b>', escapeHtml_(p.approvalLink));
  }
  return lines.join('\n');
}
