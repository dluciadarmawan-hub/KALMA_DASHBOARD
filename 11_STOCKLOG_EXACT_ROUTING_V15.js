/**
 * OWNERDASH V15 — Stock Log exact-task reader/router.
 * READ ONLY. STOCK_SUBMISSION_QUEUE remains Stock Log authority.
 * No Stock Log mutation, no Cash Bank mutation, no business logic duplication.
 */
const KOD_STOCKLOG_EXACT_V15 = Object.freeze({
  taskType: 'STOCK_REVIEW',
  sheet: 'STOCK_SUBMISSION_QUEUE',
  pendingStatuses: Object.freeze(['SUBMITTED', 'NEEDS_REVIEW', 'PENDING_REVIEW', 'NEEDS_RESOLUTION']),
  requiredHeaders: Object.freeze([
    'SUBMISSION_ID','TIMESTAMP','SUBMITTER_NAME','ROLE','ACTION_TYPE','STATUS','DATE',
    'FROM_LOCATION','TO_LOCATION','ITEM_CODE','ITEM_NAME','QTY','UNIT','NOTE'
  ])
});

function kodBuildStockLogExactTaskUrlV15_(submissionId) {
  const id = String(submissionId || '').trim();
  if (!id) return '';
  return KOD_ROUTE_LINKS.STOCK_LOG_WEBAPP
    + '?ownerTask=' + encodeURIComponent(KOD_STOCKLOG_EXACT_V15.taskType)
    + '&recordId=' + encodeURIComponent(id)
    + '&returnUrl=' + encodeURIComponent(KOD_ROUTE_LINKS.OWNER_DASHBOARD_WEBAPP);
}

function kodReadStockLogExactTasksV15_(ss) {
  const result = { rows: [], nonActionable: [], checkedSheet: '', missingHeaders: [], duplicateIds: [], scannedRows: 0 };
  const sh = ss && ss.getSheetByName(KOD_STOCKLOG_EXACT_V15.sheet);
  if (!sh) return result;
  result.checkedSheet = sh.getName();
  const lastRow = Math.min(sh.getLastRow(), KOD_SAFE_LIMITS.statusScanMaxRows);
  const lastCol = Math.min(sh.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet);
  if (lastRow < 2 || lastCol < 1) return result;
  const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const index = {};
  KOD_STOCKLOG_EXACT_V15.requiredHeaders.forEach(function(h) {
    index[h] = headers.indexOf(h);
    if (index[h] < 0) result.missingHeaders.push(h);
  });
  if (result.missingHeaders.length) return result;

  const idCounts = {};
  for (let r = 1; r < values.length; r++) {
    const id = String(values[r][index.SUBMISSION_ID] || '').trim();
    if (id) idCounts[id] = (idCounts[id] || 0) + 1;
  }
  result.duplicateIds = Object.keys(idCounts).filter(function(id) { return idCounts[id] > 1; });
  result.scannedRows = Math.max(0, values.length - 1);

  for (let r = 1; r < values.length; r++) {
    const status = String(values[r][index.STATUS] || '').trim().toUpperCase();
    if (KOD_STOCKLOG_EXACT_V15.pendingStatuses.indexOf(status) < 0) continue;
    const id = String(values[r][index.SUBMISSION_ID] || '').trim();
    const itemName = String(values[r][index.ITEM_NAME] || '').trim();
    const itemCode = String(values[r][index.ITEM_CODE] || '').trim();
    const qty = String(values[r][index.QTY] || '').trim();
    const unit = String(values[r][index.UNIT] || '').trim();
    const actionType = String(values[r][index.ACTION_TYPE] || '').trim();
    const fromLocation = String(values[r][index.FROM_LOCATION] || '').trim();
    const toLocation = String(values[r][index.TO_LOCATION] || '').trim();
    const uniqueIdentity = !!id && idCounts[id] === 1;
    const row = {
      Submission_ID: id,
      Status: status,
      Timestamp: String(values[r][index.TIMESTAMP] || '').trim(),
      Submitter_Name: String(values[r][index.SUBMITTER_NAME] || '').trim(),
      Role: String(values[r][index.ROLE] || '').trim(),
      Action_Type: actionType,
      Date: String(values[r][index.DATE] || '').trim(),
      From_Location: fromLocation,
      To_Location: toLocation,
      Item_Code: itemCode,
      Item_Name: itemName,
      Qty: qty,
      Unit: unit,
      Note: String(values[r][index.NOTE] || '').trim(),
      exactIdentity: id,
      exactIdentityType: 'SUBMISSION_ID',
      existingActionAuthority: 'EXISTING_STOCK_LOG_REVIEW_CARD',
      directExactTask: uniqueIdentity,
      actionable: uniqueIdentity,
      actionUrl: uniqueIdentity ? kodBuildStockLogExactTaskUrlV15_(id) : '',
      failureCode: id ? (uniqueIdentity ? '' : 'DUPLICATE_IDENTITY') : 'MISSING_RECORD_ID'
    };
    if (uniqueIdentity) result.rows.push(row); else result.nonActionable.push(row);
  }
  return result;
}
