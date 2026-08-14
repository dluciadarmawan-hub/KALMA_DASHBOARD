/**
 * OWNERDASH V14 — read-only payment candidate reader.
 * Source = INVOICE CUSTOMER / CUST_INV, the same persisted invoice authority used by Revenue Pending Revenue.
 * No payment amount calculation or mutation is reproduced here.
 */
const KOD_REVENUE_PAYMENT_V14 = Object.freeze({
  openRevenueStatuses:['UNPAID','PARTIAL'],
  inactiveInvoiceStatuses:[
    'CANCELLED_NO_ORDER','VOID','VOIDED','DUPLICATE_INVOICE','TEST_INVOICE','EXCLUDED_FROM_RECEIVABLE'
  ]
});

function kodRevenueHeaderIndexAnyV14_(map, names) {
  for (let i = 0; i < (names || []).length; i++) {
    if (typeof map[names[i]] === 'number') return map[names[i]];
  }
  return -1;
}

function kodRevenueValueAnyV14_(row, map, names) {
  const idx = kodRevenueHeaderIndexAnyV14_(map, names);
  return idx >= 0 ? String(row[idx] || '').trim() : '';
}

function kodRevenueMoneyDisplayToNumberV14_(value) {
  const raw = String(value || '')
    .replace(/[^0-9,.-]/g, '')
    .replace(/\./g, '')
    .replace(',', '.');
  const n = Number(raw || 0);
  return Number.isFinite(n) ? n : 0;
}

function kodReadRevenuePaymentExactTasksV13_(sheet) {
  const startRow = KOD_REVENUE_ACTION_V14.custInvDataStartRow;
  const lastRow = Math.min(
    sheet.getLastRow(),
    startRow - 1 + KOD_SAFE_LIMITS.statusScanMaxRows
  );
  const lastCol = Math.min(sheet.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet);
  if (lastRow < startRow || lastCol < 1) {
    return { rows: [], nonActionableRows: [], missingHeaders: [] };
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const hm = kodRevenueExactHeaderMapV13_(headers, [
    'NAMA CUSTOMER',
    'Revenue_Invoice_Status',
    'Revenue_Status',
    'Revenue_Total_Amount',
    'Revenue_Balance_Due',
    'Revenue_Proof_Status'
  ]);
  const invoiceIdIdx = kodRevenueHeaderIndexAnyV14_(hm.map, ['Invoice Code','Invoice_ID']);
  if (invoiceIdIdx < 0) hm.missing.push('Invoice Code|Invoice_ID');
  if (hm.missing.length) {
    return { rows: [], nonActionableRows: [], missingHeaders: hm.missing };
  }

  const rows = sheet.getRange(
    startRow,
    1,
    lastRow - startRow + 1,
    lastCol
  ).getDisplayValues();
  const byId = {};

  rows.forEach(function(row, offset) {
    const invoiceId = String(row[invoiceIdIdx] || '').trim();
    if (!invoiceId) return;
    const invoiceStatus = (
      kodRevenueExactValueV13_(row, hm.map, 'Revenue_Invoice_Status') || 'ACTIVE'
    ).toUpperCase();
    if (KOD_REVENUE_PAYMENT_V14.inactiveInvoiceStatuses.indexOf(invoiceStatus) >= 0) return;

    const revenueStatus = kodRevenueExactValueV13_(row, hm.map, 'Revenue_Status').toUpperCase();
    if (KOD_REVENUE_PAYMENT_V14.openRevenueStatuses.indexOf(revenueStatus) < 0) return;

    const balanceDue = kodRevenueExactValueV13_(row, hm.map, 'Revenue_Balance_Due');
    if (kodRevenueMoneyDisplayToNumberV14_(balanceDue) <= 0) return;

    const candidate = {
      rowNumber:startRow + offset,
      invoiceId:invoiceId,
      customerName:kodRevenueExactValueV13_(row, hm.map, 'NAMA CUSTOMER'),
      invoiceDate:kodRevenueValueAnyV14_(row, hm.map, ['Invoice Date']),
      totalAmount:kodRevenueExactValueV13_(row, hm.map, 'Revenue_Total_Amount'),
      balanceDue:balanceDue,
      revenueStatus:revenueStatus,
      proofStatus:kodRevenueExactValueV13_(row, hm.map, 'Revenue_Proof_Status')
    };
    (byId[invoiceId] || (byId[invoiceId] = [])).push(candidate);
  });

  const exactRows = [];
  const nonActionableRows = [];
  Object.keys(byId).sort().forEach(function(invoiceId) {
    const matches = byId[invoiceId];
    if (matches.length !== 1) {
      nonActionableRows.push({
        source:'Revenue',
        kind:'SOURCE_APP_QUEUE',
        taskType:KOD_REVENUE_ACTION_V14.paymentTask,
        count:1,
        title:matches[0].customerName || invoiceId,
        note:'Duplicate Invoice_ID di source Dashboard. Exact action ditahan.',
        actionLabel:'BUKA REVENUE',
        actionUrl:KOD_ROUTE_LINKS.REVENUE_WEBAPP,
        exactIdentity:invoiceId,
        directExactTask:false,
        ownerFields:{},
        details:[]
      });
      return;
    }

    const c = matches[0];
    const actionUrl = kodBuildRevenueExactTaskUrlV13_(
      KOD_REVENUE_ACTION_V14.paymentTask,
      invoiceId
    );
    if (!actionUrl) return;

    exactRows.push({
      source:'Revenue',
      kind:'ACTIONABLE_ROUTE',
      taskType:KOD_REVENUE_ACTION_V14.paymentTask,
      count:1,
      title:c.customerName || invoiceId,
      note:'Invoice ' + invoiceId + ' · sisa ' + (c.balanceDue || '-')
        + ' · ' + c.revenueStatus + ' · exact Pending Revenue action card',
      actionLabel:'PERIKSA PEMBAYARAN',
      actionUrl:actionUrl,
      exactIdentity:invoiceId,
      directExactTask:true,
      details:[],
      ownerFields:{
        Customer_Name:c.customerName,
        Invoice_ID:invoiceId,
        Invoice_Date:c.invoiceDate,
        Total_Amount:c.totalAmount,
        Balance_Due:c.balanceDue,
        Revenue_Status:c.revenueStatus,
        Proof_Status:c.proofStatus || 'BELUM ADA BUKTI'
      }
    });
  });

  return {
    rows:exactRows,
    nonActionableRows:nonActionableRows,
    missingHeaders:[]
  };
}
