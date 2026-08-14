/**
 * OWNERDASH V14 — Revenue actionable routing orchestrator.
 * READ ONLY. Payment exact authority is keyed by Invoice_ID.
 * Canteen is NON ACTIONABLE because accepted Revenue has no update-existing settlement authority.
 */
const KOD_REVENUE_ACTION_V14 = Object.freeze({
  custInvDataStartRow: 273,
  paymentTask: 'REVENUE_PAYMENT',
  canteenTask: 'CANTEEN_SETTLEMENT'
});

function kodBuildRevenueExactTaskUrlV13_(taskType, recordId) {
  const task = String(taskType || '').trim();
  const id = String(recordId || '').trim();
  if (!id || task !== KOD_REVENUE_ACTION_V14.paymentTask) return '';
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
  const missing = (required || []).filter(function(name) {
    return typeof map[name] === 'undefined';
  });
  return { map: map, missing: missing };
}

function kodRevenueExactValueV13_(row, map, name) {
  const idx = map[name];
  return typeof idx === 'number' ? String(row[idx] || '').trim() : '';
}

function kodReadRevenueCanteenExactTasksV13_(sheet) {
  const lastRow = Math.min(sheet.getLastRow(), KOD_SAFE_LIMITS.statusScanMaxRows);
  const lastCol = Math.min(sheet.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet);
  if (lastRow < 2 || lastCol < 1) return { rows: [], missingHeaders: [] };

  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const hm = kodRevenueExactHeaderMapV13_(display[0], [
    'Sales_Date', 'School', 'Source_Reference', 'Settlement_Status'
  ]);
  if (hm.missing.length) return { rows: [], missingHeaders: hm.missing };

  const allowed = {
    PENDING_SETTLEMENT:true,
    PARTIAL_SETTLED:true,
    PENDING_BUKPOT:true,
    CHECK_DIFFERENCE:true
  };
  const out = [];
  for (let r = 1; r < display.length; r++) {
    const status = kodRevenueExactValueV13_(display[r], hm.map, 'Settlement_Status').toUpperCase();
    if (!allowed[status]) continue;
    out.push({
      source:'Revenue',
      kind:'SOURCE_APP_QUEUE',
      taskType:KOD_REVENUE_ACTION_V14.canteenTask,
      count:1,
      title:kodRevenueExactValueV13_(display[r], hm.map, 'School') || 'Canteen settlement',
      note:'SOURCE ACTION AUTHORITY NOT AVAILABLE — NON ACTIONABLE · '
        + (kodRevenueExactValueV13_(display[r], hm.map, 'Sales_Date') || '-') + ' · ' + status,
      actionLabel:'BUKA REVENUE',
      actionUrl:KOD_ROUTE_LINKS.REVENUE_WEBAPP,
      exactIdentity:kodRevenueExactValueV13_(display[r], hm.map, 'Source_Reference'),
      directExactTask:false,
      ownerFields:{},
      details:[]
    });
  }
  return { rows: out, missingHeaders: [] };
}

function kodReadRevenueSummaryV13_() {
  const base = kodReadRevenueSummary_();
  if (!base || !base.ok) return base;
  try {
    const revenueSs = SpreadsheetApp.openById(KOD_SOURCE_IDS.REVENUE);
    const invoiceSs = SpreadsheetApp.openById(KOD_SOURCE_IDS.INVOICE_CUSTOMER);
    const routedRows = [];
    const schemaWarnings = [];

    const custInv = invoiceSs.getSheetByName('CUST_INV');
    if (custInv) {
      const payment = kodReadRevenuePaymentExactTasksV13_(custInv);
      routedRows.push.apply(routedRows, payment.rows || []);
      routedRows.push.apply(routedRows, payment.nonActionableRows || []);
      if ((payment.missingHeaders || []).length) {
        schemaWarnings.push('CUST_INV: ' + payment.missingHeaders.join(', '));
      }
    } else {
      schemaWarnings.push('CUST_INV: SHEET_MISSING');
    }

    const canteen = revenueSs.getSheetByName('CANTEEN_DAILY_SALES');
    if (canteen) {
      const canteenRead = kodReadRevenueCanteenExactTasksV13_(canteen);
      routedRows.push.apply(routedRows, canteenRead.rows || []);
      if ((canteenRead.missingHeaders || []).length) {
        schemaWarnings.push('CANTEEN_DAILY_SALES: ' + canteenRead.missingHeaders.join(', '));
      }
    }

    if (schemaWarnings.length) {
      base.ok = false;
      base.exactSchemaWarnings = schemaWarnings;
      base.note = 'Revenue actionable-route schema missing: ' + schemaWarnings.join(' | ');
      return base;
    }

    const remaining = (base.pendingRows || []).filter(function(r) {
      const type = String(r && r.type || '');
      return type !== 'REVENUE_INTAKE' && type !== 'CANTEEN_DAILY_SALES';
    });
    base.pendingRows = kodLimitRows_(
      routedRows.concat(remaining),
      KOD_SAFE_LIMITS.maxDashboardRowsPerPanel
    );
    base.pendingCount = Number(base.pendingRows.length || 0);
    base.exactRouteContract = {
      paymentIdentity:'Invoice_ID',
      paymentAuthority:'EXISTING_PENDING_REVENUE_INVOICE_CARD',
      paymentRoute:'REVENUE_PAYMENT',
      canteenIdentity:'Source_Reference',
      canteenAuthority:'SOURCE ACTION AUTHORITY NOT AVAILABLE — NON ACTIONABLE',
      routeMutation:false
    };
    return base;
  } catch (err) {
    base.ok = false;
    base.note = 'Revenue actionable-route read failed safely: ' + String(err && err.message || err);
    return base;
  }
}
