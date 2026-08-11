/**
 * Developer-only helpers. Keep private trailing underscore to avoid owner dropdown smoke-test dependency.
 */
function notificationEngineStaticSelfCheck_() {
  const payload = buildNotificationPayload_({
    module: 'Purchasing',
    sourceSheet: 'EXPENSE_SUBMISSION',
    recordId: 'EXP-STATIC-001',
    status: 'NEEDS_OWNER_APPROVAL',
    submittedBy: 'Rofiq',
    amount: 1250000,
    vendor: 'CV ABC',
    description: 'Pembelian bahan harian',
    approvalLink: 'https://script.google.com/example'
  });
  return {
    ok: true,
    build: KALMA_NOTIFY_CONFIG.BUILD_NAME,
    approvalRequired: isApprovalRequired_(payload.status),
    notificationKey: payload.notificationKey,
    telegramOnly: KALMA_NOTIFY_CONFIG.ACTIVE_CHANNELS.join(',') === 'Telegram',
    emailDisabled: KALMA_NOTIFY_CONFIG.EMAIL.enabled === false,
    payloadPreview: payload.messageText
  };
}
