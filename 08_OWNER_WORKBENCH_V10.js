/**
 * OWNERDASH V10 — stable-app owner workbench assembler.
 * READ ONLY. No source mutation. Source WebApps remain business authority.
 * Purchasing intentionally excluded while Purchasing is still on progress.
 */
function kodBuildStableOwnerWork_(receiving, masterApproval, stock, cash, revenue) {
  const out = [];

  (receiving && receiving.rows || []).forEach(function(r) {
    if (r.actionable !== true || !r.actionUrl) return;
    const item = String(r.Item_Name || r.Item_Code || 'Item Receiving');
    const qty = [r.Purchase_Qty || '-', r.Purchase_Unit || ''].filter(Boolean).join(' ');
    out.push({
      source: 'Receiving',
      kind: 'EXACT_TASK',
      title: item,
      note: 'Barang diterima ' + qty + ' · isi qty stok aktual dalam ' + String(r.Stock_Unit || '-'),
      count: 1,
      actionLabel: 'ISI QTY STOK AKTUAL',
      actionUrl: r.actionUrl,
      exactIdentity: r.Receiving_Line_ID || '',
      directExactTask: true,
      details: [r]
    });
  });

  if (masterApproval && Number(masterApproval.pendingCount || 0) > 0) {
    out.push({
      source: 'Order Board',
      kind: 'SOURCE_APP_QUEUE',
      title: 'Master Approval',
      note: Number(masterApproval.pendingCount || 0) + ' candidate menunggu keputusan owner',
      count: Number(masterApproval.pendingCount || 0),
      actionLabel: 'BUKA ORDER BOARD',
      actionUrl: KOD_ROUTE_LINKS.ORDER_BOARD_WEBAPP,
      directExactTask: false,
      details: (masterApproval.rows || []).slice(0, 20)
    });
  }

  (stock && stock.rows || []).filter(function(r) {
    return String(r.type || '') === 'STOCK_SUBMISSION_QUEUE';
  }).forEach(function(r) {
    out.push({
      source: 'Stock Log',
      kind: 'SOURCE_APP_QUEUE',
      title: r.title || 'Stock movement perlu review',
      note: r.note || 'Submission queue',
      count: Number(r.count || 0),
      actionLabel: 'BUKA STOCK LOG',
      actionUrl: KOD_ROUTE_LINKS.STOCK_LOG_WEBAPP,
      directExactTask: false,
      details: r.details || []
    });
  });

  (cash && cash.pendingRows || []).forEach(function(r) {
    out.push({
      source: 'Cash Bank',
      kind: 'SOURCE_APP_QUEUE',
      title: r.title || 'Payment request',
      note: r.note || 'Pending payment request',
      count: Number(r.count || 0),
      actionLabel: 'BUKA CASH BANK',
      actionUrl: KOD_ROUTE_LINKS.CASH_BANK_WEBAPP,
      directExactTask: false,
      details: r.details || []
    });
  });

  (revenue && revenue.pendingRows || []).forEach(function(r) {
    out.push({
      source: 'Revenue',
      kind: 'SOURCE_APP_QUEUE',
      title: r.title || 'Revenue perlu dicek',
      note: r.note || r.type || 'Pending revenue work',
      count: Number(r.count || 0),
      actionLabel: 'BUKA REVENUE',
      actionUrl: KOD_ROUTE_LINKS.REVENUE_WEBAPP,
      directExactTask: false,
      details: r.details || []
    });
  });

  return out.filter(function(r) {
    return r.count > 0 || r.directExactTask === true;
  }).slice(0, 18);
}

function kodBuildOwnerBriefV10_(result) {
  const items = [];
  const ownerWork = result.ownerWork || [];

  const cashAccounts = result.cashSnapshot || [];
  const cashById = {};
  cashAccounts.forEach(function(a) { cashById[String(a.account || '').toUpperCase()] = a; });
  const bca = cashById.BCA_KALMA;
  const mandiri = cashById.MANDIRI_CECE_QRIS;
  const pettyKalma = cashById.PETTYCASH_KALMA;
  const pettyCanteen = cashById.PETTYCASH_CANTEEN;
  const keyCash = [bca, mandiri, pettyKalma, pettyCanteen].filter(Boolean);
  if (keyCash.length) {
    const primary = bca || keyCash[0];
    const rest = keyCash.filter(function(a) { return a !== primary; }).map(function(a) {
      return String(a.accountName || a.account || 'Account') + ' ' + String(a.balance || 'Rp0');
    });
    const pendingTotal = keyCash.reduce(function(sum, a) { return sum + (Number(a.pendingOut) || 0); }, 0);
    const hasCashWarning = keyCash.some(function(a) { return !!a.warning; });
    items.push({
      tone: hasCashWarning ? 'warn' : 'ok',
      title: 'Saldo terakhir · ' + String(primary.accountName || primary.account || 'Cash Bank') + ' ' + String(primary.balance || 'Rp0'),
      note: (rest.length ? rest.join(' · ') : 'Cash Bank terbaca') + (pendingTotal > 0 ? ' · ada pending pembayaran' : '')
    });
  }
  const exactReceiving = ownerWork.filter(function(r) {
    return r.source === 'Receiving' && r.directExactTask === true;
  });
  if (exactReceiving.length) {
    items.push({
      tone: 'danger',
      title: exactReceiving.length + ' tugas Receiving bisa langsung dibuka',
      note: 'Exact Receiving_Line_ID sudah terhubung ke task stock count.'
    });
  }

  const groupedCount = ownerWork.filter(function(r) {
    return r.source !== 'Receiving';
  }).reduce(function(sum, r) {
    return sum + (Number(r.count) || 0);
  }, 0);
  if (groupedCount) {
    items.push({
      tone: 'warn',
      title: groupedCount + ' pekerjaan dari app stabil menunggu',
      note: 'Order Board · Stock Log · Cash Bank · Revenue. Satu klik membuka source app yang benar.'
    });
  }

  const warnings = result.warnings || [];
  if (warnings.length) {
    items.push({
      tone: 'danger',
      title: warnings.length + ' peringatan penting',
      note: warnings[0].source + ' · ' + warnings[0].title
    });
  }

  const prodRows = result.production || [];
  const realProd = prodRows.filter(function(r) { return r.pack !== 'ACTION'; });
  if (!realProd.length) {
    items.push({
      tone: 'warn',
      title: 'Produksi besok belum terbaca aman',
      note: 'Menu/pack besok belum terbaca sebagai data produksi aktif.'
    });
  }

  if (!items.length) {
    items.push({
      tone: 'ok',
      title: 'Tidak ada pekerjaan owner utama terbaca',
      note: 'Queue app stabil kosong dan tidak ada peringatan penting.'
    });
  }
  return items.slice(0, 6);
}
