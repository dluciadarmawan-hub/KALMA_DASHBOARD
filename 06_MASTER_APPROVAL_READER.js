/**
 * OWNERDASH_20260627_V04A — Pending Master Approval reader.
 * READ ONLY. Reads Order Board INBOUND_MASTER_APPROVAL_QUEUE only.
 * Supports Order Board V25 queue headers and V24 source fields:
 * - First_Order_ID
 * - First_Order_Line_ID
 * - Source_Order_IDs
 * - Source_Order_Line_IDs
 * No writes to VENDOR_REGISTRY, MASTER_BAHAN, PRICE_MASTER, PRICE_UPDATES,
 * Receiving, Stock Log, Purchasing, Cash Bank, finance, ledger, or any row delete.
 */
function kodReadOrderBoardMasterApprovalSummary_() {
  const startedAt = new Date();
  const requiredV25Headers = [
    'Candidate_ID', 'Candidate_Type', 'Candidate_Status', 'Raw_Typed_Name', 'Normalized_Name',
    'First_Order_ID', 'First_Order_Line_ID', 'Source_Order_IDs', 'Source_Order_Line_IDs',
    'Vendor_Candidate_ID', 'Item_Candidate_ID', 'Vendor_Master_ID', 'Vendor_Resolved_Name',
    'Master_Item_ID', 'Item_Resolved_Name', 'Item_Resolved_Unit',
    'Created_At', 'Created_By', 'Resolved_At', 'Resolved_By'
  ];
  const out = {
    source: 'ORDER_BOARD_MASTER_APPROVAL',
    ok: false,
    note: 'Not loaded yet',
    checkedSheets: [KOD_MASTER_APPROVAL_R1.queueSheet],
    refreshedAt: startedAt.toISOString(),
    pendingCount: 0,
    vendorCount: 0,
    itemCount: 0,
    rows: [],
    pendingRows: [],
    warnings: [],
    details: [],
    storageProof: {
      spreadsheetId: KOD_SOURCE_IDS.ORDER_BOARD,
      spreadsheetTitle: '',
      expectedSheet: KOD_MASTER_APPROVAL_R1.queueSheet,
      sheetExists: false,
      requiredV25Headers: requiredV25Headers,
      observedHeaders: [],
      missingRequiredHeaders: requiredV25Headers.slice(),
      rowCount: 0,
      dataRowCount: 0,
      pendingRowCount: 0,
      isQueueEmpty: true,
      proofStatus: 'NOT_CHECKED'
    },
    sourceMap: {
      sourceSpreadsheetId: KOD_SOURCE_IDS.ORDER_BOARD,
      sourceSpreadsheetTitle: '',
      sourceSheet: KOD_MASTER_APPROVAL_R1.queueSheet,
      sourceBuild: KOD_MASTER_APPROVAL_R1.sourceBuild,
      sourceSha256: KOD_MASTER_APPROVAL_R1.sourceSha256,
      supportedOrderBoardV25Fields: [
        'First_Order_ID', 'First_Order_Line_ID', 'Source_Order_IDs', 'Source_Order_Line_IDs'
      ],
      writePolicy: 'READ_ONLY_NO_MASTER_OR_DOWNSTREAM_MUTATION'
    }
  };

  try {
    const ss = SpreadsheetApp.openById(KOD_SOURCE_IDS.ORDER_BOARD);
    const spreadsheetTitle = ss.getName ? ss.getName() : '';
    out.storageProof.spreadsheetTitle = spreadsheetTitle;
    out.sourceMap.sourceSpreadsheetTitle = spreadsheetTitle;

    const sh = ss.getSheetByName(KOD_MASTER_APPROVAL_R1.queueSheet);
    out.storageProof.sheetExists = !!sh;
    if (!sh) {
      out.note = 'Sheet missing: ' + KOD_MASTER_APPROVAL_R1.queueSheet;
      out.storageProof.proofStatus = 'SHEET_MISSING';
      out.warnings.push({
        source: 'ORDER_BOARD',
        title: 'Master approval queue belum ditemukan',
        note: 'Sheet ' + KOD_MASTER_APPROVAL_R1.queueSheet + ' belum ada di Order Board storage.',
        details: [out.storageProof, out.sourceMap]
      });
      return out;
    }

    const values = sh.getDataRange().getDisplayValues();
    out.storageProof.rowCount = values ? values.length : 0;
    out.storageProof.dataRowCount = values && values.length ? Math.max(values.length - 1, 0) : 0;
    out.storageProof.isQueueEmpty = !values || values.length < 2;
    if (!values || !values.length) {
      out.ok = true;
      out.note = 'Queue sheet exists, but has no header or rows. Sheet exists: YES · Row count: 0.';
      out.storageProof.proofStatus = 'SHEET_EXISTS_EMPTY';
      out.warnings.push({
        source: 'ORDER_BOARD',
        title: 'Master approval queue kosong',
        note: 'Sheet ada, tapi header belum terbaca.',
        details: [out.storageProof]
      });
      return out;
    }

    const observedHeaders = (values[0] || []).map(function(h) { return String(h || '').trim(); }).filter(Boolean);
    out.storageProof.observedHeaders = observedHeaders;
    const headerMap = kodMasterApprovalHeaderMap_(values[0]);
    const missingRequired = requiredV25Headers.filter(function(h) {
      return headerMap[kodMasterApprovalNorm_(h)] == null;
    });
    out.storageProof.missingRequiredHeaders = missingRequired;
    out.storageProof.proofStatus = missingRequired.length ? 'HEADER_MISMATCH' : 'V25_HEADER_MATCH';

    if (missingRequired.length) {
      out.warnings.push({
        source: 'ORDER_BOARD',
        title: 'Order Board V25 header mismatch',
        note: 'Missing headers: ' + missingRequired.join(', '),
        details: [out.storageProof]
      });
    }

    if (values.length < 2) {
      out.ok = missingRequired.length === 0;
      out.note = 'Queue sheet exists. Row count: ' + out.storageProof.dataRowCount + ' data rows. Pending rows: 0.';
      return out;
    }

    const closedRe = new RegExp(KOD_MASTER_APPROVAL_R1.closedStatusRegex, 'i');
    const rows = [];

    for (let r = 1; r < values.length; r++) {
      const source = kodMasterApprovalRowObject_(values[r], headerMap);
      const candidateId = kodMasterApprovalFirst_(source, ['Candidate_ID', 'Candidate Id', 'CandidateID', 'candidate_id']);
      const status = kodMasterApprovalFirst_(source, ['Status', 'Candidate_Status', 'Approval_Status']);
      const rawTypedName = kodMasterApprovalRawName_(source);
      const candidateType = kodMasterApprovalInferType_(source);

      if (!candidateId && !rawTypedName) continue;
      if (closedRe.test(String(status || '').trim())) continue;

      const firstOrderId = kodMasterApprovalFirst_(source, [
        'First_Order_ID', 'First Order ID', 'FirstOrderID',
        'Source_Order_ID', 'Source Order ID', 'Order_ID', 'Order ID'
      ]);
      const firstLineId = kodMasterApprovalFirst_(source, [
        'First_Order_Line_ID', 'First Order Line ID', 'FirstOrderLineID',
        'Source_Order_Line_ID', 'Source_Order_LineID', 'Source_Line_ID', 'Source Line ID', 'Line_ID', 'Line ID', 'Order_Line_ID'
      ]);
      const allOrderIds = kodMasterApprovalFirst_(source, [
        'Source_Order_IDs', 'Source Order IDs', 'All_Source_Order_IDs', 'All Source Order IDs',
        'Source_Order_ID', 'Source Order ID', 'Order_ID', 'Order ID'
      ]);
      const allLineIds = kodMasterApprovalFirst_(source, [
        'Source_Order_Line_IDs', 'Source Order Line IDs', 'All_Source_Order_Line_IDs', 'All Source Order Line IDs',
        'Source_Line_IDs', 'Source Line IDs', 'Source_Order_Line_ID', 'Source_Order_LineID',
        'Source_Line_ID', 'Source Line ID', 'Order_Line_ID', 'Line_ID', 'Line ID'
      ]);

      const item = {
        Candidate_ID: candidateId || ('ROW-' + (r + 1)),
        Candidate_Type: candidateType,
        Raw_Typed_Name: rawTypedName,
        Source_Module: kodMasterApprovalFirst_(source, ['Source_Module', 'Source Module', 'Module']) || 'ORDER_BOARD',
        First_Order_ID: firstOrderId,
        First_Order_Line_ID: firstLineId,
        Source_Order_IDs: allOrderIds,
        Source_Order_Line_IDs: allLineIds,
        Source_Order_ID: firstOrderId || kodMasterApprovalFirst_(source, ['Source_Order_ID', 'Source Order ID', 'Order_ID', 'Order ID']),
        Source_Line_ID: firstLineId || kodMasterApprovalFirst_(source, ['Source_Line_ID', 'Source Line ID', 'Line_ID', 'Line ID', 'Order_Line_ID']),
        Status: status || 'PENDING',
        Created_At: kodMasterApprovalFirst_(source, ['Created_At', 'Created At', 'Timestamp', 'Created']),
        Suggested_Action: kodMasterApprovalSuggestedAction_(candidateType),
        actionButtonsDisabled: true,
        rowNumber: r + 1,
        sourceBuild: KOD_MASTER_APPROVAL_R1.sourceBuild,
        sourceSha256: KOD_MASTER_APPROVAL_R1.sourceSha256,
        storageProofStatus: out.storageProof.proofStatus
      };
      rows.push(item);
    }

    out.ok = missingRequired.length === 0;
    out.rows = rows.slice(0, 80);
    out.pendingCount = rows.length;
    out.storageProof.pendingRowCount = rows.length;
    out.storageProof.isQueueEmpty = rows.length === 0;
    out.vendorCount = rows.filter(function(x) { return x.Candidate_Type === 'VENDOR'; }).length;
    out.itemCount = rows.filter(function(x) { return x.Candidate_Type === 'ITEM'; }).length;
    out.note = rows.length ? ('Queue sheet exists. Row count: ' + out.storageProof.dataRowCount + ' data rows. Pending rows: ' + rows.length + '.') : ('Queue sheet exists. Row count: ' + out.storageProof.dataRowCount + ' data rows. Pending rows: 0.');
    if (missingRequired.length) out.note = 'Read with V25 header mismatch: ' + missingRequired.join(', ');
    out.sourceMap.storageProof = out.storageProof;
    out.pendingRows = rows.slice(0, KOD_SAFE_LIMITS.maxDashboardRowsPerPanel).map(function(x) {
      return {
        source: 'ORDER_BOARD',
        title: x.Candidate_Type + ' · ' + (x.Raw_Typed_Name || x.Candidate_ID),
        note: 'Candidate ' + x.Candidate_ID + ' · ' + x.Status + ' · First Order ' + (x.First_Order_ID || '-'),
        count: 'PENDING',
        details: [x]
      };
    });
    out.details = rows;
    return out;
  } catch (err) {
    out.note = 'Read failed safely.';
    out.error = kodMasterApprovalSafeError_(err);
    out.storageProof.proofStatus = 'READ_FAILED_SAFE';
    out.warnings.push({
      source: 'ORDER_BOARD',
      title: 'Master approval queue read failed',
      note: out.error,
      details: [out.storageProof, out.sourceMap]
    });
    return out;
  }
}

function kodMasterApprovalHeaderMap_(headers) {
  const map = {};
  (headers || []).forEach(function(h, i) {
    const key = kodMasterApprovalNorm_(h);
    if (key && map[key] == null) map[key] = i;
  });
  return map;
}

function kodMasterApprovalRowObject_(row, headerMap) {
  const obj = {};
  Object.keys(headerMap || {}).forEach(function(norm) {
    obj[norm] = row[headerMap[norm]] || '';
  });
  return obj;
}

function kodMasterApprovalFirst_(obj, names) {
  for (let i = 0; i < names.length; i++) {
    const v = obj[kodMasterApprovalNorm_(names[i])];
    if (v != null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function kodMasterApprovalNorm_(v) {
  return String(v || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function kodMasterApprovalInferType_(source) {
  const explicit = kodMasterApprovalFirst_(source, ['Candidate_Type', 'Type', 'Master_Type', 'Entity_Type']).toUpperCase();
  if (/VENDOR|SUPPLIER/.test(explicit)) return 'VENDOR';
  if (/ITEM|BAHAN|MATERIAL|PRODUCT/.test(explicit)) return 'ITEM';

  const id = kodMasterApprovalFirst_(source, ['Candidate_ID', 'Candidate Id', 'CandidateID']).toUpperCase();
  if (/VENDOR|SUPPLIER|VEN|VND/.test(id)) return 'VENDOR';
  if (/ITEM|BAHAN|ITM|MAT/.test(id)) return 'ITEM';

  const raw = kodMasterApprovalRawName_(source).toUpperCase();
  if (/PT |CV |TOKO|DEPOT|SUPPLIER|VENDOR/.test(raw)) return 'VENDOR';
  return 'ITEM';
}

function kodMasterApprovalRawName_(source) {
  return kodMasterApprovalFirst_(source, [
    'Raw_Typed_Name', 'Raw Typed Name', 'Typed_Name', 'Typed Name', 'Raw_Name',
    'Candidate_Name', 'Candidate Name', 'Vendor_Name', 'Vendor Name', 'Item_Name',
    'Item Name', 'Nama', 'Nama_Bahan', 'Nama Vendor', 'Nama Bahan'
  ]);
}

function kodMasterApprovalSuggestedAction_(type) {
  return type === 'VENDOR' ? 'Review vendor candidate — action disabled in R1' : 'Review item candidate — action disabled in R1';
}

function kodMasterApprovalSafeError_(err) {
  const msg = err && err.message ? err.message : String(err || 'Unknown error');
  return msg.replace(/https?:\/\/\S+/g, '[url]').slice(0, 240);
}
