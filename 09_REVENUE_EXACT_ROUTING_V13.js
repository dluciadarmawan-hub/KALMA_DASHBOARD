/**
 * OWNERDASH V13 — bounded Revenue exact-task reader + route builder.
 * READ ONLY. It does not modify Revenue or Cash Bank.
 * Isolated from 02_SOURCE_READERS.js so all existing V12 readers stay byte-exact.
 */
function kodBuildRevenueExactTaskUrlV13_(taskType, recordId) {
  const task = String(taskType || '').trim();
  const id = String(recordId || '').trim();
  if (!id || ['REVENUE_PAYMENT', 'CANTEEN_SETTLEMENT'].indexOf(task) < 0) return '';
  return KOD_ROUTE_LINKS.REVENUE_WEBAPP
    + '?ownerTask=' + encodeURIComponent(task)
    + '&recordId=' + encodeURIComponent(id)
    + '&returnUrl=' + encodeURIComponent(KOD_ROUTE_LINKS.OWNER_DASHBOARD_WEBAPP);
}

function kodRevenueExactHeaderMapV13_(headers, required) {
  const map = {};
  (headers || []).forEach(function(h, idx) {
    const key = String(h || '').trim();
    if (key && typeof map[key] === 'undefined') map[key] = idx;
  });
  const missing = (required || []).filter(function(name) { return typeof map[name] === 'undefined'; });
  return { map: map, missing: missing };
}

function kodRevenueExactValueV13_(row, map, name) {
  const idx = map[name];
  return typeof idx === 'number' ? String(row[idx] || '').trim() : '';
}

function kodReadRevenuePaymentExactTasksV13_(sheet) {
  const lastRow = Math.min(sheet.getLastRow(), KOD_SAFE_LIMITS.statusScanMaxRows);
  const lastCol = Math.min(sheet.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet);
  if (lastRow < 2 || lastCol < 1) return { rows: [], missingHeaders: [] };
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const required = ['Revenue_ID','Payment_Date','Customer_Name','Paid_Amount','Received_Into','Payment_Status'];
  const hm = kodRevenueExactHeaderMapV13_(display[0], required);
  if (hm.missing.length) return { rows: [], missingHeaders: hm.missing };
  const allowed = { PENDING_CHECK:true, NEEDS_CHECK:true, MISMATCH:true };
  const out = [];
  for (let r = 1; r < display.length; r++) {
    const status = kodRevenueExactValueV13_(display[r], hm.map, 'Payment_Status').toUpperCase();
    if (!allowed[status]) continue;
    const id = kodRevenueExactValueV13_(display[r], hm.map, 'Revenue_ID');
    if (!id) continue;
    const actionUrl = kodBuildRevenueExactTaskUrlV13_('REVENUE_PAYMENT', id);
    if (!actionUrl) continue;
    out.push({
      source:'Revenue', kind:'EXACT_TASK', taskType:'REVENUE_PAYMENT', count:1,
      title:kodRevenueExactValueV13_(display[r], hm.map, 'Customer_Name') || id,
      note:'Pembayaran exact dari REVENUE_INTAKE', actionLabel:'PERIKSA PEMBAYARAN', actionUrl:actionUrl,
      exactIdentity:id, directExactTask:true, details:[],
      ownerFields:{
        Customer_Name:kodRevenueExactValueV13_(display[r], hm.map, 'Customer_Name'),
        Payment_Date:kodRevenueExactValueV13_(display[r], hm.map, 'Payment_Date'),
        Paid_Amount:kodRevenueExactValueV13_(display[r], hm.map, 'Paid_Amount'),
        Received_Into:kodRevenueExactValueV13_(display[r], hm.map, 'Received_Into'),
        Payment_Status:kodRevenueExactValueV13_(display[r], hm.map, 'Payment_Status'),
        Revenue_ID:id
      }
    });
  }
  return { rows: out, missingHeaders: [] };
}

function kodReadRevenueCanteenExactTasksV13_(sheet) {
  const lastRow = Math.min(sheet.getLastRow(), KOD_SAFE_LIMITS.statusScanMaxRows);
  const lastCol = Math.min(sheet.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet);
  if (lastRow < 2 || lastCol < 1) return { rows: [], missingHeaders: [] };
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const required = ['Sales_Date','School','Total_Sales','Source_Reference','Net_Transfer_Expected','Settlement_Amount_Received','Settlement_Status'];
  const hm = kodRevenueExactHeaderMapV13_(display[0], required);
  if (hm.missing.length) return { rows: [], missingHeaders: hm.missing };
  const allowed = { PENDING_SETTLEMENT:true, PARTIAL_SETTLED:true, PENDING_BUKPOT:true, CHECK_DIFFERENCE:true };
  const out = [];
  for (let r = 1; r < display.length; r++) {
    const status = kodRevenueExactValueV13_(display[r], hm.map, 'Settlement_Status').toUpperCase();
    if (!allowed[status]) continue;
    const id = kodRevenueExactValueV13_(display[r], hm.map, 'Source_Reference');
    if (!id) continue;
    const actionUrl = kodBuildRevenueExactTaskUrlV13_('CANTEEN_SETTLEMENT', id);
    if (!actionUrl) continue;
    out.push({
      source:'Revenue', kind:'EXACT_TASK', taskType:'CANTEEN_SETTLEMENT', count:1,
      title:kodRevenueExactValueV13_(display[r], hm.map, 'School') || id,
      note:'Settlement exact dari CANTEEN_DAILY_SALES', actionLabel:'PERIKSA SETTLEMENT', actionUrl:actionUrl,
      exactIdentity:id, directExactTask:true, details:[],
      ownerFields:{
        School:kodRevenueExactValueV13_(display[r], hm.map, 'School'),
        Sales_Date:kodRevenueExactValueV13_(display[r], hm.map, 'Sales_Date'),
        Total_Sales:kodRevenueExactValueV13_(display[r], hm.map, 'Total_Sales'),
        Settlement_Status:kodRevenueExactValueV13_(display[r], hm.map, 'Settlement_Status'),
        Net_Transfer_Expected:kodRevenueExactValueV13_(display[r], hm.map, 'Net_Transfer_Expected'),
        Settlement_Amount_Received:kodRevenueExactValueV13_(display[r], hm.map, 'Settlement_Amount_Received'),
        Source_Reference:id
      }
    });
  }
  return { rows: out, missingHeaders: [] };
}

function kodReadRevenueSummaryV13_() {
  const base = kodReadRevenueSummary_();
  if (!base || !base.ok) return base;
  try {
    const ss = SpreadsheetApp.openById(KOD_SOURCE_IDS.REVENUE);
    const exactRows = [];
    const schemaWarnings = [];
    const intake = ss.getSheetByName('REVENUE_INTAKE');
    if (intake) {
      const x = kodReadRevenuePaymentExactTasksV13_(intake);
      exactRows.push.apply(exactRows, x.rows || []);
      if ((x.missingHeaders || []).length) schemaWarnings.push('REVENUE_INTAKE: ' + x.missingHeaders.join(', '));
    }
    const canteen = ss.getSheetByName('CANTEEN_DAILY_SALES');
    if (canteen) {
      const x = kodReadRevenueCanteenExactTasksV13_(canteen);
      exactRows.push.apply(exactRows, x.rows || []);
      if ((x.missingHeaders || []).length) schemaWarnings.push('CANTEEN_DAILY_SALES: ' + x.missingHeaders.join(', '));
    }
    if (schemaWarnings.length) {
      base.ok = false;
      base.exactSchemaWarnings = schemaWarnings;
      base.note = 'Exact Revenue route schema missing: ' + schemaWarnings.join(' | ');
      return base;
    }
    const remaining = (base.pendingRows || []).filter(function(r) {
      return String(r && r.type || '') !== 'REVENUE_INTAKE' && String(r && r.type || '') !== 'CANTEEN_DAILY_SALES';
    });
    base.pendingRows = kodLimitRows_(exactRows.concat(remaining), KOD_SAFE_LIMITS.maxDashboardRowsPerPanel);
    base.pendingCount = exactRows.length + remaining.reduce(function(sum, r) { return sum + (Number(r && r.count) || 0); }, 0);
    base.exactTaskCount = exactRows.length;
    base.note = 'Connected. Revenue exact tasks use canonical Revenue_ID / Source_Reference; other Revenue queues remain source-app routes.';
    return base;
  } catch (err) {
    return kodFail_(base, err);
  }
}
