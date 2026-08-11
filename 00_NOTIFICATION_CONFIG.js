/**
 * KALMA OWNER NOTIFICATION ENGINE V1.0
 * Shared infrastructure for Kalma WebApp approval notifications.
 * Telegram is the only active transport in this release.
 * Email / WhatsApp / Google Chat providers are intentionally disabled.
 */
const KALMA_NOTIFY_CONFIG = Object.freeze({
  BUILD_NAME: 'KALMA_OWNER_NOTIFICATION_ENGINE_V1_0',
  RELEASE_CHANNEL: 'TELEGRAM_ONLY',
  DEFAULT_TIMEZONE: 'Asia/Jakarta',
  NOTIFICATION_LOG_SHEET: 'NOTIFICATION_LOG',
  ACTIVE_CHANNELS: Object.freeze(['Telegram']),
  DISABLED_CHANNELS: Object.freeze(['Email', 'WhatsApp', 'GoogleChat']),
  TELEGRAM: Object.freeze({
    enabled: true,
    chatId: '8977960828',
    tokenScriptProperty: 'KALMA_TELEGRAM_BOT_TOKEN',
    parseMode: 'HTML',
    apiBaseUrl: 'https://api.telegram.org/bot'
  }),
  EMAIL: Object.freeze({ enabled: false }),
  WHATSAPP: Object.freeze({ enabled: false }),
  GOOGLE_CHAT: Object.freeze({ enabled: false }),
  APPROVAL_STATUSES: Object.freeze([
    'NEEDS_OWNER_APPROVAL',
    'NEEDS_OWNER_REVIEW',
    'WAITING_OWNER_VERIFY',
    'NEEDS_PAYMENT_APPROVAL',
    'NEEDS_BCA_CHECK',
    'NEEDS_CASH_APPROVAL',
    'NEEDS_PAYROLL_APPROVAL'
  ]),
  FINAL_STATUSES: Object.freeze([
    'APPROVED',
    'PAID',
    'POSTED',
    'SYNCED',
    'SYNCED_CLEAN',
    'CLOSED',
    'VOIDED',
    'CANCELLED',
    'REJECTED',
    'DONE',
    'COMPLETED'
  ]),
  LOG_HEADERS: Object.freeze([
    'Timestamp',
    'Module',
    'Source Sheet',
    'Record ID',
    'Status',
    'Submitted By',
    'Amount',
    'Vendor',
    'Description',
    'Notification Channel',
    'Notification Result',
    'Notification Key',
    'Spreadsheet URL',
    'Error Message'
  ])
});
