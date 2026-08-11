function cleanText_(value) {
  return String(value == null ? '' : value).replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function cleanKeyPart_(value) {
  return cleanText_(value).replace(/\|/g, '/').toUpperCase();
}

function normalizeAmount_(value) {
  if (typeof value === 'number') return value;
  const text = String(value == null ? '' : value).replace(/[^0-9,-]/g, '').replace(/\./g, '').replace(',', '.');
  const n = Number(text);
  return Number.isFinite(n) ? n : 0;
}

function formatRupiah_(value) {
  const n = normalizeAmount_(value);
  return 'Rp ' + Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function formatTimestamp_(date) {
  const d = date instanceof Date ? date : new Date(date);
  return Utilities.formatDate(d, KALMA_NOTIFY_CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function escapeHtml_(text) {
  return String(text == null ? '' : text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getActiveSpreadsheetUrlSafe_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss ? ss.getUrl() : '';
  } catch (_) {
    return '';
  }
}
