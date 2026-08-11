/** Public read-only endpoint used by ui.html. */
function kodGetOwnerDashboardV0(payload) {
  const startedAt = new Date();
  const authSession = kodRequireDashboardSession_(payload || {}, KOD_AUTH.roleSets.dashboardRead, 'kodGetOwnerDashboardV0');
  const result = {
    ok: true,
    build: KOD_BUILD,
    currentUser: kodPublicSession_(authSession),
    auth: kodPublicSession_(authSession),
    generatedAt: startedAt.toISOString(),
    timezone: Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Asia/Jakarta',
    cards: [],
    ownerBrief: [],
    production: [],
    staffTasks: [],
    pending: [],
    warnings: [],
    speedDial: kodBuildSpeedDial_(),
    sourceHealth: [],
    cashSnapshot: [],
    stockSnapshot: [],
    invoiceSnapshot: [],
    receivingTasks: [],
    receivingTaskMeta: {},
    drilldowns: {},
    notes: [
      'Dashboard V1 is read-only.',
      'No posting, no sync, no ledger mutation.',
      'Dashboard reads source systems directly and shows action summary first.'
    ]
  };

  const authGate = kodCheckDashboardAuthorization_();
  result.authGate = authGate;
  if (!authGate.ok) {
    const blocked = ['MASTER_SCHEDULE', 'PURCHASING', 'REVENUE', 'CASH_BANK', 'STOCK_LOG', 'INVOICE_CUSTOMER', 'RECEIVING'].map(function(src) {
      return { source: src, ok: false, note: 'AUTH_GATE_PENDING', checkedSheets: [] };
    });
    result.sourceHealth = blocked;
    result.cards = kodBuildAuthGateCards_(authGate);
    result.ownerBrief = [{ tone: 'danger', title: 'Authorization belum aktif', note: authGate.safeMessage }];
    result.warnings = [{ source: 'Setup', title: 'Spreadsheet authorization belum aktif', note: authGate.safeMessage }];
    result.notes.push('AUTH_GATE_PENDING: source readers paused to avoid raw SpreadsheetApp errors in UI.');
    return result;
  }

  const master = kodReadMasterScheduleSummary_();
  const purchasing = kodReadPurchasingSummary_();
  const revenue = kodReadRevenueSummary_();
  const cash = kodReadCashBankSummary_();
  const stock = kodReadStockSummary_();
  const invoice = kodReadInvoiceSummary_();
  const masterApproval = kodReadOrderBoardMasterApprovalSummary_();
  const receiving = kodReadReceivingStockCountTasks_();

  [master, purchasing, revenue, cash, stock, invoice, masterApproval, receiving].forEach(function(x) {
    result.sourceHealth.push({ source: x.source, ok: x.ok, note: x.note || '', checkedSheets: x.checkedSheets || [], refreshedAt: x.refreshedAt || result.generatedAt });
  });

  const pendingCount = (purchasing.pendingCount || 0) + (revenue.pendingCount || 0) + (cash.pendingRows || []).reduce(function(sum, x) { return sum + (Number(x.count) || 0); }, 0);
  const receivingActionCount = (receiving.rows || []).filter(function(r) { return r.actionable === true; }).length;
  const alertCount = (cash.warningCount || 0) + (stock.warningCount || 0);
  const hasProductionAction = (master.productionRows || []).some(function(r) { return r.pack === 'ACTION'; });

  result.cards = [
    { key: 'receiving_stock_count', title: 'Receiving — Qty Stok Aktual', value: String(receivingActionCount), note: 'Exact deep-link task dari RECEIVING_STOCK_QTY_COUNT_TASK', tone: receivingActionCount ? 'danger' : (receiving.ok ? 'ok' : 'warn'), icon: '', drillKey: 'receiving' },
    { key: 'menu_besok', title: 'Menu Besok', value: master.menuCountText || '—', note: master.note || 'Master Schedule reader ready', tone: hasProductionAction ? 'warn' : (master.ok ? 'ok' : 'warn'), icon: '', drillKey: 'production' },
    { key: 'pending_owner', title: 'Pending Source App', value: String(pendingCount), note: 'Purchasing + Revenue + Cash. Dikerjakan di source WebApp masing-masing.', tone: pendingCount ? 'warn' : 'ok', icon: '', drillKey: 'pending' },
    { key: 'pending_master_approval', title: 'Order Board Pending', value: String(masterApproval.pendingCount || 0), note: 'Source app required. Belum actionable dari Dashboard.', tone: (masterApproval.pendingCount || 0) ? 'warn' : (masterApproval.ok ? 'ok' : 'warn'), icon: '', drillKey: 'masterApproval' },
    { key: 'alert_merah', title: 'Peringatan Penting', value: String(alertCount), note: 'Cash + stock warning. Invoice tampil terpisah.', tone: alertCount ? 'danger' : 'ok', icon: '', drillKey: 'warnings' }
  ];

  result.production = master.productionRows || [];
  result.pendingMasterApproval = masterApproval.rows || [];
  result.pendingMasterApprovalMeta = masterApproval.sourceMap || {};
  result.pendingMasterApprovalStorageProof = masterApproval.storageProof || {};
  result.pending = [].concat(purchasing.pendingRows || [], revenue.pendingRows || [], cash.pendingRows || []);
  result.receivingTasks = receiving.rows || [];
  result.receivingTaskMeta = receiving.meta || {};
  result.staffTasks = [].concat(stock.staffTasks || []);
  result.warnings = [].concat(masterApproval.warnings || [], cash.warnings || [], stock.warnings || [], invoice.warnings || []);
  result.cashSnapshot = cash.accounts || [];
  result.stockSnapshot = stock.rows || [];
  result.invoiceSnapshot = invoice.rows || [];
  result.drilldowns = kodBuildDrilldowns_(result, master, purchasing, revenue, cash, stock, invoice, masterApproval, receiving);
  result.ownerBrief = kodBuildOwnerBrief_(result);
  return result;
}

function kodBuildDrilldowns_(result, master, purchasing, revenue, cash, stock, invoice, masterApproval, receiving) {
  return {
    production: kodFlattenDetails_(result.production || [], master.productionDetails || []),
    pending: kodFlattenDetails_(result.pending || []),
    masterApproval: kodFlattenDetails_(result.pendingMasterApproval || [], (masterApproval && masterApproval.details) || []),
    warnings: kodFlattenDetails_(result.warnings || []),
    staff: kodFlattenDetails_(result.staffTasks || []),
    cash: kodFlattenDetails_(result.cashSnapshot || []),
    stock: kodFlattenDetails_(result.stockSnapshot || []),
    invoice: kodFlattenDetails_(result.invoiceSnapshot || []),
    receiving: kodFlattenDetails_(result.receivingTasks || [], (receiving && receiving.details) || [])
  };
}

function kodFlattenDetails_(rows, fallback) {
  const details = [];
  (rows || []).forEach(function(r) {
    if (r.details && r.details.length) {
      r.details.forEach(function(d) { details.push(d); });
    } else {
      details.push(r);
    }
  });
  if (!details.length && fallback && fallback.length) {
    fallback.forEach(function(f) { details.push(f); });
  }
  return details.slice(0, 60);
}

function kodCheckDashboardAuthorization_() {
  const out = {
    ok: false,
    status: 'AUTH_GATE_PENDING',
    safeMessage: 'Dashboard belum punya authorization aktif untuk membaca spreadsheet sumber.',
    detail: ''
  };
  try {
    const ss = SpreadsheetApp.openById(KOD_SOURCE_IDS.APPS_SCRIPT_PROJECTS);
    out.ok = true;
    out.status = 'AUTH_OK';
    out.safeMessage = 'Dashboard authorization OK.';
    out.detail = ss.getName ? ss.getName() : 'Spreadsheet opened';
    return out;
  } catch (err) {
    out.detail = err && err.message ? err.message : String(err);
    if (/permission|authorization|auth|scope/i.test(out.detail)) {
      out.safeMessage = 'Authorization spreadsheet belum aktif. Source readers ditahan dulu supaya dashboard tidak menampilkan error mentah.';
    } else {
      out.safeMessage = 'Dashboard belum bisa membaca spreadsheet registry. Cek access file atau Script ID.';
    }
    return out;
  }
}

/** Manual authorization runner. READ ONLY. */
function kodAuthorizeOwnerDashboard(payload) {
  const authSession = kodRequireOwnerDashboardSession_(payload || {}, 'kodAuthorizeOwnerDashboard');
  const probe = SpreadsheetApp.openById(KOD_SOURCE_IDS.APPS_SCRIPT_PROJECTS);
  return {
    ok: true,
    status: 'AUTH_OK',
    source: 'APPS_SCRIPT_PROJECTS',
    spreadsheetName: probe.getName(),
    noWritesPerformed: true,
    auth: kodPublicSession_(authSession),
    instruction: 'Authorization active. Refresh WebApp dashboard.'
  };
}

function kodBuildAuthGateCards_(authGate) {
  return [
    { key: 'auth_gate', title: 'Setup Source Access', value: 'AUTH', note: authGate.safeMessage, tone: 'warn', icon: 'alert' },
    { key: 'menu_besok', title: 'Menu Besok', value: '—', note: 'Ditahan sampai authorization OK', tone: 'neutral', icon: 'tray' },
    { key: 'pending_owner', title: 'Pending Owner', value: '—', note: 'Ditahan sampai source reader aktif', tone: 'neutral', icon: 'pending' },
    { key: 'alert_merah', title: 'Alert Merah', value: '—', note: 'Ditahan sampai source reader aktif', tone: 'neutral', icon: 'alert' }
  ];
}

function kodBuildSpeedDial_() {
  return [
    { label: 'Receiving', subtitle: 'Penerimaan + task qty stok aktual', url: KOD_ROUTE_LINKS.RECEIVING_WEBAPP, enabled: true, group: 'Operations' },
    { label: 'Order Board', subtitle: 'Order + master approval source', url: KOD_ROUTE_LINKS.ORDER_BOARD_WEBAPP, enabled: true, group: 'Operations' },
    { label: 'Master Schedule', subtitle: 'Menu, pack, KJS', url: KOD_ROUTE_LINKS.MASTER_SCHEDULE_WEBAPP, enabled: true, group: 'Operations' },
    { label: 'Purchasing', subtitle: 'Nota, item, approval', url: KOD_ROUTE_LINKS.PURCHASING_WEBAPP, enabled: true, group: 'Finance' },
    { label: 'Revenue', subtitle: 'Customer bayar, sales', url: KOD_ROUTE_LINKS.REVENUE_WEBAPP, enabled: true, group: 'Finance' },
    { label: 'Cash Bank', subtitle: 'Saldo, cashbox, QRIS', url: KOD_ROUTE_LINKS.CASH_BANK_WEBAPP, enabled: true, group: 'Finance' },
    { label: 'Stock Log', subtitle: 'Stok kritis, movement', url: KOD_ROUTE_LINKS.STOCK_LOG_WEBAPP, enabled: true, group: 'Stock' },
    { label: 'Invoice Customer', subtitle: 'Invoice, paid/unpaid', url: KOD_ROUTE_LINKS.INVOICE_WEBAPP, enabled: true, group: 'Billing' },
    { label: 'Payroll', subtitle: 'Payroll owner', url: KOD_ROUTE_LINKS.PAYROLL_WEBAPP, enabled: true, group: 'Payroll' },
    { label: 'Owner Dashboard', subtitle: 'Dashboard ini', url: KOD_ROUTE_LINKS.OWNER_DASHBOARD_WEBAPP, enabled: true, group: 'Owner' }
  ];
}
