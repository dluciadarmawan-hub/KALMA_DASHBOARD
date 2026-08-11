function doGet(e) {
  return HtmlService.createTemplateFromFile('ui')
    .evaluate()
    .setTitle('KALMA OWNER DASHBOARD')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function kodPing() {
  return {
    ok: true,
    build: KOD_BUILD,
    timestamp: new Date().toISOString(),
    mode: 'standalone-dashboard-read-only-central-auth-required',
    authRequiredForDashboard: true
  };
}
