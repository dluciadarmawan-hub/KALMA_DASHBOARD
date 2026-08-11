/**
 * KALMA OWNER DASHBOARD source readers.
 * V1.6: trusted readers. Staff task source is Stock Log TASK_SUBMISSION / daily daftar tugas. Cash snapshot follows Cash Bank ledger.
 * READ ONLY. Uses SpreadsheetApp.openById + getRange + getValues/getDisplayValues only.
 */
function kodReadMasterScheduleSummary_() {
  const out = kodSourceOut_('MASTER_SCHEDULE');
  try {
    const ss = SpreadsheetApp.openById(KOD_SOURCE_IDS.MASTER_SCHEDULE);
    const sheetNames = ss.getSheets().map(function(s) { return s.getName(); });
    out.ok = true;
    out.checkedSheets = sheetNames.slice(0, 18);

    const tomorrowKey = kodDateOffsetKey_(1);
    const independent = kodBuildIndependentProductionSnapshot_(ss, tomorrowKey);
    const kcs = independent.rows.length ? { rows: [], totalPack: 0, checkedSheets: [], note: '' } : kodReadKcsCookSummary_(ss, tomorrowKey);
    const finalRows = independent.rows.length ? independent.rows : kcs.rows;
    const totalPack = independent.rows.length ? independent.totalPack : kcs.totalPack;

    out.checkedSheets = kodUnique_([].concat(independent.checkedSheets || [], kcs.checkedSheets || [], out.checkedSheets.slice(0, 8)));
    out.menuCountText = finalRows.length ? (finalRows.length + ' menu') : 'Belum tampil';
    out.packText = totalPack ? (kodFormatNumber_(totalPack) + ' porsi/pack') : 'Belum terkunci';
    out.productionRows = finalRows;
    out.productionDetails = independent.details || [];
    out.tomorrowKey = tomorrowKey;
    out.sourceMode = independent.rows.length ? 'INDEPENDENT_MASTER_SCHEDULE_READER' : (kcs.rows.length ? 'KCS_FALLBACK' : 'NO_PRODUCTION_ROW');
    out.note = independent.rows.length ? independent.note : (kcs.note || 'Master Schedule terbaca, tapi produksi besok belum bisa dihitung otomatis.');
    out.actionNeeded = '';

    if (!finalRows.length) {
      out.actionNeeded = 'Master Schedule reader tidak menemukan row besok. Cek tanggal/status/header order.';
      out.productionRows = [{
        program: 'MASTER SCHEDULE',
        menu: 'Belum ada order/menu aktif untuk besok yang bisa dibaca dashboard.',
        pack: 'ACTION',
        note: 'Tanggal target: ' + tomorrowKey + '. Dashboard tidak mengarang data produksi.'
      }];
    }
    return out;
  } catch (err) {
    return kodFail_(out, err);
  }
}

function kodBuildIndependentProductionSnapshot_(ss, tomorrowKey) {
  const out = { rows: [], totalPack: 0, checkedSheets: [], details: [], note: '' };
  let rows = [];

  // V1.3 exact readers only. No raw row guessing.
  rows = rows.concat(kodReadDailyWideProductionV13_(ss, tomorrowKey));
  rows = rows.concat(kodReadKidsTeenWideProductionV13_(ss, kodMonthSheetNameV13_('KIDS', tomorrowKey), 'Kids Catering', tomorrowKey));
  rows = rows.concat(kodReadKidsTeenWideProductionV13_(ss, kodMonthSheetNameV13_('TEEN', tomorrowKey), 'Teen Catering', tomorrowKey));
  rows = rows.concat(kodReadHealthyWideProductionV13_(ss, kodMonthSheetNameV13_('HLT', tomorrowKey), tomorrowKey));
  rows = rows.concat(kodReadOrderCountProductionV13_(ss, tomorrowKey));

  const grouped = {};
  rows.forEach(function(r) {
    if (!kodIsSafeProductionRow_(r, tomorrowKey)) return;
    const pack = kodSafeProductionPack_(r.packNum, r.packNum, r.menu);
    if (!pack) return;
    const key = kodNormalizeKey_((r.program || '') + '|' + (r.menu || '') + '|' + (r.meal || ''));
    if (!key) return;
    if (!grouped[key]) {
      grouped[key] = { program: r.program || 'Program', menu: r.menu, packNum: 0, note: r.note || '', meal: r.meal || '', details: [], sourceSheets: [] };
    }
    grouped[key].packNum += pack;
    grouped[key].details.push(r.detail || r);
    if (r.sourceSheet) {
      grouped[key].sourceSheets.push(r.sourceSheet);
      out.checkedSheets.push(r.sourceSheet);
    }
  });

  let finalRows = Object.keys(grouped).map(function(k) {
    const g = grouped[k];
    const title = g.meal ? (g.program + ' · ' + g.meal) : g.program;
    return {
      program: title,
      menu: g.menu,
      pack: kodFormatNumber_(g.packNum),
      packNum: g.packNum,
      note: (g.note || 'Reader exact') + (g.sourceSheets.length ? ' · ' + kodUnique_(g.sourceSheets).join(', ') : ''),
      details: g.details.slice(0, 12)
    };
  }).sort(function(a, b) { return (b.packNum || 0) - (a.packNum || 0); });

  const totalPack = finalRows.reduce(function(sum, r) { return sum + (r.packNum || 0); }, 0);
  if (totalPack > (KOD_SAFE_LIMITS.maxProductionTotalPack || 15000)) {
    out.rows = [];
    out.totalPack = 0;
    out.checkedSheets = kodUnique_(out.checkedSheets);
    out.details = finalRows;
    out.note = 'Production reader diblokir: total pack tidak masuk akal. Dashboard tidak menampilkan angka palsu.';
    return out;
  }

  out.totalPack = totalPack;
  out.rows = kodLimitRows_(finalRows, KOD_SAFE_LIMITS.maxDashboardRowsPerPanel || 8);
  out.details = finalRows;
  out.checkedSheets = kodUnique_(out.checkedSheets);
  out.note = finalRows.length ? ('Produksi besok dibaca dari exact Master Schedule readers: ' + out.checkedSheets.join(', ')) : 'Exact production reader tidak menemukan produksi besok.';
  return out;
}


function kodReadDailyWideProductionV12_(ss, tomorrowKey) {
  return kodReadDailyWideProductionV13_(ss, tomorrowKey);
}

function kodReadDailyWideProductionV13_(ss, tomorrowKey) {
  const out = [];
  const sh = kodFindSheetByNamesV12_(ss, [kodMonthSheetNameV13_('DAILY', tomorrowKey), kodShortMonthSheetNameV13_('D', tomorrowKey)]);
  if (!sh) return out;
  const lastCol = sh.getLastColumn();
  const lastRow = Math.min(sh.getLastRow(), 30);
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  for (let c = 0; c < (values[0] || []).length; c++) {
    if (kodDateKey_(values[0][c]) !== tomorrowKey) continue;
    const menuParts = [display[1] && display[1][c], display[3] && display[3][c], display[5] && display[5][c]].filter(function(x) { return String(x || '').trim(); });
    const menu = kodCleanMenuV13_(menuParts.join(' / '));
    const pax = kodSafeProductionPack_(values[9] && values[9][c + 1], display[9] && display[9][c + 1], menu);
    if (!menu || !pax) continue;
    out.push({
      sourceSheet: sh.getName(),
      program: 'Daily Catering',
      meal: '',
      menu: menu,
      packNum: pax,
      note: 'Daily total pax',
      detail: { sheet: sh.getName(), row: 10, column: c + 2, date: tomorrowKey, menu: menu, qty: pax, source: 'DAILY_TOTAL_PAX' }
    });
  }
  return out;
}


function kodReadKidsTeenWideProductionV12_(ss, sheetName, program, tomorrowKey) {
  return kodReadKidsTeenWideProductionV13_(ss, sheetName, program, tomorrowKey);
}

function kodReadKidsTeenWideProductionV13_(ss, sheetName, program, tomorrowKey) {
  const out = [];
  const sh = ss.getSheetByName(sheetName) || ss.getSheetByName(kodShortMonthSheetNameV13_(program.indexOf('Teen') >= 0 ? 'T' : 'K', tomorrowKey));
  if (!sh) return out;
  const lastCol = sh.getLastColumn();
  const lastRow = Math.min(sh.getLastRow(), 80);
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  for (let c = 0; c < (values[1] || []).length; c++) {
    if (kodDateKey_(values[1][c]) !== tomorrowKey) continue;
    const menu = kodCleanMenuV13_(display[2] && display[2][c]);
    const total = kodSafeProductionPack_(values[0] && values[0][c + 1], display[0] && display[0][c + 1], menu);
    if (!menu || !total) continue;
    out.push({
      sourceSheet: sh.getName(),
      program: program,
      meal: '',
      menu: menu,
      packNum: total,
      note: 'Total porsi dari header harian',
      detail: { sheet: sh.getName(), row: 1, column: c + 2, date: tomorrowKey, menu: menu, qty: total, source: 'KIDS_TEEN_TOTAL_HEADER' }
    });
  }
  return out;
}


function kodReadHealthyWideProductionV12_(ss, sheetName, tomorrowKey) {
  return kodReadHealthyWideProductionV13_(ss, sheetName, tomorrowKey);
}

function kodReadHealthyWideProductionV13_(ss, sheetName, tomorrowKey) {
  const out = [];
  const sh = ss.getSheetByName(sheetName) || ss.getSheetByName(kodShortMonthSheetNameV13_('H', tomorrowKey));
  if (!sh) return out;
  const lastCol = sh.getLastColumn();
  const lastRow = Math.min(sh.getLastRow(), 42);
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  for (let c = 0; c < (values[2] || []).length; c++) {
    if (kodDateKey_(values[2][c]) !== tomorrowKey) continue;
    const lunchMenu = kodCleanMenuV13_(display[3] && display[3][c]);
    const dinnerMenu = kodCleanMenuV13_(display[5] && display[5][c]);
    const lunchQty = kodSafeProductionPack_(values[8] && values[8][c + 1], display[8] && display[8][c + 1], lunchMenu);
    const dinnerQty = kodSafeProductionPack_(values[8] && values[8][c + 5], display[8] && display[8][c + 5], dinnerMenu);
    if (lunchMenu && lunchQty) out.push({ sourceSheet: sh.getName(), program: 'Healthy Catering', meal: 'Lunch', menu: lunchMenu, packNum: lunchQty, note: 'Healthy lunch pax', detail: { sheet: sh.getName(), row: 9, column: c + 2, date: tomorrowKey, menu: lunchMenu, qty: lunchQty, source: 'HLT_LUNCH' } });
    if (dinnerMenu && dinnerQty) out.push({ sourceSheet: sh.getName(), program: 'Healthy Catering', meal: 'Dinner', menu: dinnerMenu, packNum: dinnerQty, note: 'Healthy dinner pax', detail: { sheet: sh.getName(), row: 9, column: c + 6, date: tomorrowKey, menu: dinnerMenu, qty: dinnerQty, source: 'HLT_DINNER' } });
  }
  return out;
}


function kodReadOrderCountProductionV12_(ss, tomorrowKey) {
  return kodReadOrderCountProductionV13_(ss, tomorrowKey);
}

function kodReadOrderCountProductionV13_(ss, tomorrowKey) {
  // V1.3: ORDER_COUNT is customer/order list, not menu production source.
  // It is not used for menu totals to avoid mixing date range rows into production cards.
  return [];
}


function kodOrderCountDateMatchesV12_(value, tomorrowKey) {
  if (!value) return false;
  const direct = kodDateKey_(value);
  if (direct === tomorrowKey) return true;
  const s = String(value || '');
  const target = kodDateFromKeyV12_(tomorrowKey);
  if (!target) return false;
  const range = s.match(/([A-Za-z]{3})\s*(\d{1,2})\s*[–-]\s*([A-Za-z]{3})?\s*(\d{1,2}),?\s*(20\d{2})/i);
  if (range) {
    const m1 = kodMonthIndexV12_(range[1]);
    const d1 = Number(range[2]);
    const m2 = kodMonthIndexV12_(range[3] || range[1]);
    const d2 = Number(range[4]);
    const y = Number(range[5]);
    const a = new Date(y, m1, d1);
    const b = new Date(y, m2, d2);
    return target >= a && target <= b;
  }
  return false;
}


function kodMonthSheetNameV13_(prefix, key) {
  const d = kodDateFromKeyV12_(key) || new Date();
  const y = d.getFullYear();
  const m = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'][d.getMonth()];
  return String(prefix || '').toUpperCase() + '_' + m + '_' + y;
}
function kodShortMonthSheetNameV13_(prefix, key) {
  const d = kodDateFromKeyV12_(key) || new Date();
  const y = String(d.getFullYear());
  const m = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'][d.getMonth()];
  return String(prefix || '').toUpperCase() + '_' + m + '_' + y;
}
function kodCleanMenuV13_(text) {
  text = String(text || '').replace(/\s+/g, ' ').trim();
  if (!text) return '';
  if (/format stiker|list kurir|apartemen|unit\s+\d+|double protein|tanpa karbo/i.test(text)) return '';
  if (/\b(mon|tue|wed|thu|fri|sat|sun|senin|selasa|rabu|kamis|jumat|sabtu|minggu)\b/i.test(text)) return '';
  if (/\b20\d{2}\b/.test(text) && text.length < 60) return '';
  return kodShort_(text, 120);
}

function kodDateFromKeyV12_(key) { const m = String(key || '').match(/^(20\d{2})-(\d{2})-(\d{2})$/); return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null; }
function kodMonthIndexV12_(m) { const map = {JAN:0,FEB:1,MAR:2,APR:3,MAY:4,MEI:4,JUN:5,JUL:6,AUG:7,AGU:7,SEP:8,OCT:9,OKT:9,NOV:10,DEC:11,DES:11}; return map[String(m || '').slice(0,3).toUpperCase()] || 0; }
function kodCleanMenuV12_(text) { text = String(text || '').replace(/\s+/g, ' ').trim(); if (!text || kodLooksLikeProductionNoise_(text)) return ''; return kodShort_(text, 120); }
function kodFindSheetByNamesV12_(ss, names) { for (let i = 0; i < names.length; i++) { const sh = ss.getSheetByName(names[i]); if (sh) return sh; } return null; }

function kodProductionCandidateSheetNames_(ss) {
  const all = ss.getSheets().map(function(s) { return s.getName(); });
  const now = new Date();
  const tz = Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Asia/Jakarta';
  const y = Utilities.formatDate(now, tz, 'yyyy');
  const yy = Utilities.formatDate(now, tz, 'yy');
  const mm = Utilities.formatDate(now, tz, 'MM');
  const monthShort = Utilities.formatDate(now, tz, 'MMM').toUpperCase();
  const preferred = [
    'ORDERS_LOG', 'ORDER_COUNT_' + y + '_' + mm,
    'DAILY_' + monthShort + '_' + y, 'DAILY_JUN_2026', 'DAILY_MAY_2026',
    'HLT_' + monthShort + '_' + y, 'HLT_JUN_2026', 'HLT_MENU_ALL',
    'KIDS_' + monthShort + '_' + y, 'KIDS_JUN_2026',
    'TEEN_' + monthShort + '_' + y, 'TEEN_JUN_2026',
    'MENU_GABUNGAN2', 'MENU_GABUNGAN',
    'D_' + monthShort + '_' + yy, 'D_JUN_2026',
    'K_' + monthShort + '_' + yy, 'K_JUN_2026',
    'T_' + monthShort + '_' + yy, 'T_JUN_2026',
    'H_' + monthShort + '_' + yy, 'H_JUN_2026',
    'BIGSTEPS', 'SAMPOERNA PI'
  ];
  const found = [];
  preferred.forEach(function(p) {
    all.forEach(function(name) {
      if (kodNormalizeKey_(name) === kodNormalizeKey_(p) && found.indexOf(name) < 0) found.push(name);
    });
  });
  all.forEach(function(name) {
    if (found.length >= 28) return;
    if (/^(ORDERS|ORDER_COUNT|DAILY|HLT|KIDS|TEEN|MENU_GABUNGAN|D_|K_|T_|H_|BIGSTEPS|SAMPOERNA)/i.test(name) && found.indexOf(name) < 0) found.push(name);
  });
  return found.slice(0, 28);
}

function kodProbeSheetForProduction_(sheet, tomorrowKey) {
  const out = { rows: [] };
  const lastRow = Math.min(sheet.getLastRow(), KOD_SAFE_LIMITS.maxRowsPerSheet || 900);
  const lastCol = Math.min(sheet.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet || 40);
  if (lastRow < 2 || lastCol < 2) return out;
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const tabular = kodProbeTabularProduction_(sheet.getName(), values, display, tomorrowKey);
  // V1.1 safety: wide date grids are NOT trusted until exact contract is locked.
  // They caused date serial / date-range text to be counted as production pack.
  out.rows = [].concat(tabular.rows || []);
  return out;
}

function kodProbeTabularProduction_(sheetName, values, display, tomorrowKey) {
  const out = { rows: [] };
  const headerRow = kodFindHeaderRow_(display, ['date', 'tanggal', 'tgl', 'menu', 'program', 'customer', 'size', 'porsi', 'qty', 'status']);
  if (headerRow < 0) return out;
  const header = display[headerRow].map(function(h) { return String(h || '').trim().toLowerCase(); });
  const idxDate = kodFirstHeader_(header, ['delivery date', 'tanggal kirim', 'tgl kirim', 'tanggal', 'date', 'tgl', 'hari']);
  const idxMenu = kodFirstHeader_(header, ['menu name', 'menu utama', 'main menu', 'menu', 'makanan']);
  const idxProgram = kodFirstHeader_(header, ['program', 'product', 'category', 'kategori', 'package', 'paket']);
  const idxQty = kodFirstHeader_(header, ['total porsi', 'porsi masak', 'porsi', 'pack', 'qty', 'quantity', 'pax', 'jumlah']);
  const idxSize = kodFirstHeader_(header, ['size', 'package size', 'paket', 'portion']);
  const idxStatus = kodFirstHeader_(header, ['status', 'state', 'order status']);
  const idxCustomer = kodFirstHeader_(header, ['customer', 'nama', 'name']);
  const idxNote = kodFirstHeader_(header, ['note', 'notes', 'catatan', 'special', 'allergy', 'alergi']);
  if (idxDate < 0 && idxMenu < 0) return out;

  for (let r = headerRow + 1; r < values.length; r++) {
    const row = display[r] || [];
    const statusText = idxStatus >= 0 ? row[idxStatus] : '';
    if (kodIsCancelledOrFinalProduction_(statusText)) continue;
    const rowDateKey = idxDate >= 0 ? (kodDateKey_(values[r][idxDate]) || kodDateKey_(row[idxDate])) : kodRowContainsDateKey_(row, tomorrowKey);
    if (rowDateKey !== tomorrowKey) continue;
    const menu = idxMenu >= 0 ? String(row[idxMenu] || '').trim() : kodGuessMenuFromRow_(row);
    if (!menu || kodLooksLikeNoise_(menu)) continue;
    const program = (idxProgram >= 0 ? String(row[idxProgram] || '').trim() : '') || kodProgramFromSheet_(sheetName);
    const rawQty = idxQty >= 0 ? values[r][idxQty] : '';
    const displayQty = idxQty >= 0 ? row[idxQty] : '';
    const qty = idxQty >= 0 ? kodSafeProductionPack_(rawQty, displayQty, menu) : 0;
    const sizeText = idxSize >= 0 ? row[idxSize] : '';
    const guessedQty = qty ? 0 : kodGuessQuantityFromRow_(values[r], row);
    const packNum = qty ? kodSafeProductionPack_(kodComputePackFromQtySize_(qty, sizeText), displayQty, menu) : guessedQty;
    if (!packNum) continue;
    const noteParts = [];
    if (idxCustomer >= 0 && row[idxCustomer]) noteParts.push('Customer: ' + row[idxCustomer]);
    if (idxNote >= 0 && row[idxNote]) noteParts.push(row[idxNote]);
    out.rows.push({
      sourceSheet: sheetName,
      program: program,
      menu: menu,
      packNum: packNum,
      note: noteParts.length ? kodShort_(noteParts.join(' · '), 110) : ('Besok · ' + tomorrowKey),
      confidence: 'TABULAR_DATE_MATCH',
      detail: { sheet: sheetName, row: r + 1, date: tomorrowKey, menu: menu, qty: packNum || qty || '', status: statusText || '', note: noteParts.join(' · ') }
    });
    if (out.rows.length >= 60) break;
  }
  return out;
}

function kodProbeWideDateProduction_(sheetName, values, display, tomorrowKey) {
  // Disabled in V1.1.
  // Reason: monthly/wide sheets may contain Google date serials or date-range text in data cells.
  // Dashboard must not count those as pack/porsi. Only explicit tabular date+qty rows are trusted.
  return { rows: [], disabled: true, reason: 'WIDE_DATE_READER_DISABLED_UNTIL_CONTRACT_LOCKED' };
}

function kodReadKcsCookSummary_(ss, tomorrowKey) {
  const out = { rows: [], totalPack: 0, checkedSheets: [], note: '', actionNeeded: '' };
  const sh = ss.getSheetByName('KCS_COOK_SUMMARY') || ss.getSheetByName('KCS_JUN_2026') || ss.getSheetByName('KCS_COOK_DETAIL');
  if (!sh) {
    out.note = 'KCS fallback belum ditemukan.';
    out.actionNeeded = 'Independent reader juga gagal; cek Master Schedule raw sheets.';
    return out;
  }
  out.checkedSheets.push(sh.getName());
  const lastRow = Math.min(sh.getLastRow(), KOD_SAFE_LIMITS.statusScanMaxRows);
  const lastCol = Math.min(sh.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet);
  if (lastRow < 2 || lastCol < 3) {
    out.note = 'KCS fallback kosong.';
    return out;
  }
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headerRow = kodFindHeaderRow_(display, ['date', 'tanggal', 'menu', 'porsi', 'program']);
  const hr = headerRow >= 0 ? headerRow : 0;
  const header = display[hr].map(function(h) { return String(h || '').trim().toLowerCase(); });
  const idxDate = kodFirstHeader_(header, ['date', 'tanggal']);
  const idxProgram = kodFirstHeader_(header, ['program']);
  const idxMeal = kodFirstHeader_(header, ['meal', 'shift']);
  const idxMenu = kodFirstHeader_(header, ['menuname', 'menu name', 'menu']);
  const idxPorsi = kodFirstHeader_(header, ['porsimasak', 'porsi masak', 'total porsi', 'porsi', 'qty']);
  if (idxDate < 0 || idxMenu < 0) {
    out.note = 'KCS fallback header belum cocok.';
    return out;
  }
  for (let r = hr + 1; r < values.length; r++) {
    const dateKey = kodDateKey_(values[r][idxDate]) || kodDateKey_(display[r][idxDate]);
    if (dateKey !== tomorrowKey) continue;
    const menu = String(display[r][idxMenu] || '').trim();
    if (!menu) continue;
    const program = String(display[r][idxProgram >= 0 ? idxProgram : 1] || '').trim() || 'PROGRAM';
    const meal = String(display[r][idxMeal >= 0 ? idxMeal : 2] || '').trim();
    const pack = kodAsNumber_(values[r][idxPorsi >= 0 ? idxPorsi : -1] || display[r][idxPorsi >= 0 ? idxPorsi : -1]);
    out.totalPack += pack || 0;
    out.rows.push({ program: program + (meal ? ' · ' + meal : ''), menu: menu, pack: pack ? kodFormatNumber_(pack) : 'cek', packNum: pack || 0, note: 'Fallback KCS · ' + tomorrowKey, details: [{ sheet: sh.getName(), row: r + 1, menu: menu, qty: pack || '' }] });
    if (out.rows.length >= KOD_SAFE_LIMITS.maxDashboardRowsPerPanel) break;
  }
  out.note = out.rows.length ? 'Fallback KCS dipakai.' : 'KCS fallback tidak punya row besok.';
  return out;
}

function kodReadPurchasingSummary_() {
  const out = kodSourceOut_('PURCHASING');
  try {
    const ss = SpreadsheetApp.openById(KOD_SOURCE_IDS.PURCHASING);
    const groups = [
      // V1.4: TASK_QUEUE is not an owner-pending source. It can contain stale staff/history rows
      // even when Approval Center and My Task are already empty. Do not count it in Owner Dashboard.
      { sheet: 'BCA_PAYMENT_CHECKLIST', title: 'Bukti BCA perlu owner verify', note: 'Proof staff menunggu cek Cece', needles: ['POSTING_TO_VERIFY', 'WAITING_VERIFY', 'NEEDS_OWNER_VERIFY'], exclude: ['VERIFIED', 'BCA_VERIFIED', 'POSTED', 'VOIDED', 'REJECTED_FINAL', 'CLOSED'] },
      { sheet: 'SUPPLIER_PAYMENT_QUEUE', title: 'Bayar supplier perlu action', note: 'Supplier/payment queue belum final', needles: ['NEEDS_PAYMENT', 'NEEDS_APPROVAL', 'PENDING_APPROVAL', 'WAITING_PAYMENT', 'UNPAID', 'OUTSTANDING'], exclude: ['SYNCED_CLEAN', 'SYNCED CLEAN', 'PAID', 'READY_TO_PAY', 'NOT_FINAL_OR_PENDING', 'CLOSED'] },
      { sheet: 'ADVANCE_BELANJA', title: 'Advance belanja belum settle', note: 'Uang staff belum ditutup nota/refund', needles: ['OUTSTANDING', 'PARTIAL_SETTLED', 'NEEDS_SETTLEMENT'], exclude: ['SETTLED', 'CLOSED', 'VOIDED'] },
      { sheet: 'CASH_MOVEMENT_QUEUE', title: 'Mutasi/kas perlu approval', note: 'Cash movement belum final', needles: ['PENDING', 'NEEDS_APPROVAL', 'WAITING'], exclude: ['ACTIVE', 'APPROVED', 'POSTED', 'SYNCED', 'SYNCED_CLEAN', 'SYNCED CLEAN', 'VOIDED', 'CLOSED'] },
      { sheet: 'EXPENSE_SUBMISSION', title: 'Expense/kas kecil perlu review', note: 'Kas kecil/expense belum final', needles: ['RETURNED_TO_STAFF', 'NEEDS_REVIEW', 'NEEDS_APPROVAL', 'PENDING_APPROVAL', 'PENDING_FIX'], exclude: ['SYNC_NOT_REQUESTED', 'APPROVED', 'PAID', 'SYNCED', 'SYNCED_CLEAN', 'SYNCED CLEAN', 'VOIDED', 'CLOSED'] }
    ];
    let total = 0;
    const rows = [];
    let staffTasks = [];
    groups.forEach(function(g) {
      const sh = ss.getSheetByName(g.sheet);
      if (!sh) return;
      out.checkedSheets.push(g.sheet);
      const c = kodCountActionRows_(sh, g.needles, g.exclude || []);
      total += c.count;
      if (c.count) rows.push({ source: 'Purchasing', type: g.sheet, title: g.title, count: c.count, note: g.note + (c.sampleStatus ? ' · contoh: ' + kodShort_(c.sampleStatus, 78) : ''), details: c.details || [] });
    });
    out.ok = true;
    out.pendingCount = total;
    out.pendingRows = kodLimitRows_(kodSortActionRows_(rows), KOD_SAFE_LIMITS.maxDashboardRowsPerPanel);
    out.staffTasks = kodLimitRows_(kodMergeStaffTasks_(staffTasks), KOD_SAFE_LIMITS.maxDashboardRowsPerPanel);
    out.note = out.checkedSheets.length ? 'Connected. Owner-action groups only.' : 'Connected, but target queue tabs not found.';
    return out;
  } catch (err) {
    return kodFail_(out, err);
  }
}

function kodReadRevenueSummary_() {
  const out = kodSourceOut_('REVENUE');
  try {
    const ss = SpreadsheetApp.openById(KOD_SOURCE_IDS.REVENUE);
    const groups = [
      { sheet: 'REVENUE_INTAKE', title: 'Customer bayar perlu dicek', note: 'Payment intake belum clear', needles: ['PENDING_CHECK', 'NEEDS_CHECK', 'MISMATCH'], exclude: ['EXCLUDED_TEST', 'POSTED', 'SYNCED_CLEAN', 'SYNCED CLEAN', 'VOIDED', 'CLOSED'] },
      { sheet: 'CUSTOMER_PAYMENT_LOG', title: 'Payment customer mismatch/partial', note: 'Customer payment log', needles: ['PARTIAL', 'MISMATCH', 'NEEDS_CHECK'], exclude: ['PAID_FULL', 'FULLY_PAID', 'POSTED', 'SYNCED_CLEAN', 'CLOSED'] },
      { sheet: 'RECEIVABLE_TRACKER', title: 'Invoice/piutang belum lunas', note: 'Receivable tracker', needles: ['UNPAID', 'PARTIAL', 'OVERDUE'], exclude: ['PAID_FULL', 'FULLY_PAID', 'CLOSED'] },
      { sheet: 'QRIS_SETTLEMENT', title: 'QRIS belum settle', note: 'QRIS settlement', needles: ['PENDING', 'WAITING', 'UNSETTLED'], exclude: ['SETTLED', 'POSTED', 'CLOSED'] },
      { sheet: 'CANTEEN_DAILY_SALES', title: 'Sales kantin belum close', note: 'Canteen sales daily close', needles: ['OPEN', 'PENDING', 'NEEDS_CLOSE'], exclude: ['CLOSED', 'POSTED', 'SETTLED'] },
      { sheet: 'CASH_DEPOSIT_LOG', title: 'Setoran cash perlu cek', note: 'Cash deposit log', needles: ['PENDING', 'WAITING', 'NEEDS'], exclude: ['POSTED', 'CLOSED'] }
    ];
    let total = 0;
    const rows = [];
    groups.forEach(function(g) {
      const sh = ss.getSheetByName(g.sheet);
      if (!sh) return;
      out.checkedSheets.push(g.sheet);
      const c = kodCountActionRows_(sh, g.needles, g.exclude || []);
      total += c.count;
      if (c.count) rows.push({ source: 'Revenue', type: g.sheet, title: g.title, count: c.count, note: g.note + (c.sampleStatus ? ' · contoh: ' + kodShort_(c.sampleStatus, 78) : ''), details: c.details || [] });
    });
    out.ok = true;
    out.pendingCount = total;
    out.pendingRows = kodLimitRows_(kodSortActionRows_(rows), KOD_SAFE_LIMITS.maxDashboardRowsPerPanel);
    out.note = out.checkedSheets.length ? 'Connected. Revenue action groups only.' : 'Connected, but revenue queue tabs not found.';
    return out;
  } catch (err) {
    return kodFail_(out, err);
  }
}

function kodReadCashBankSummary_() {
  const out = kodSourceOut_('CASH_BANK');
  try {
    const ss = SpreadsheetApp.openById(KOD_SOURCE_IDS.CASH_BANK);
    const snapshot = kodReadComputedCashSnapshot_(ss);
    const paymentQueue = ss.getSheetByName('PAYMENT_REQUEST_QUEUE');
    const pending = paymentQueue ? kodCountActionRows_(paymentQueue, ['NEEDS_PAYMENT', 'READY_TO_PAY', 'WAITING_PAYMENT'], ['PAID', 'POSTED', 'VOIDED', 'CLOSED']) : { count: 0, details: [] };
    out.ok = true;
    out.checkedSheets = kodUnique_(['ACCOUNT_MASTER'].concat(snapshot.checkedSheets || []).concat(paymentQueue ? ['PAYMENT_REQUEST_QUEUE'] : []));
    out.accounts = snapshot.accounts;
    out.cashSnapshotMeta = snapshot.meta;
    out.warningCount = snapshot.accounts.filter(function(a) { return a.warning; }).length + pending.count;
    out.pendingRows = pending.count ? [{ source: 'Cash Bank', type: 'PAYMENT_REQUEST_QUEUE', title: 'Payment request belum dieksekusi', count: pending.count, note: pending.sampleStatus ? 'contoh: ' + kodShort_(pending.sampleStatus, 78) : 'Pending payment request', details: pending.details || [] }] : [];
    out.warnings = snapshot.accounts.filter(function(a) { return a.warning; }).map(function(a) { return { source: 'Cash Bank', title: a.account || 'Account', note: a.warning, details: [{ account: a.account, balance: a.balance, formula: a.note }] }; });
    out.note = 'Connected. Cash snapshot = opening balance + ledger delta + cashbox delta.';
    return out;
  } catch (err) {
    return kodFail_(out, err);
  }
}

function kodReadStockSummary_() {
  const out = kodSourceOut_('STOCK_LOG');
  try {
    const ss = SpreadsheetApp.openById(KOD_SOURCE_IDS.STOCK_LOG);
    const groups = [
      { sheet: 'STOK_TERBARU', title: 'Stok kritis / minus', note: 'Stok terbaru', mode: 'stock_level' },
      { sheet: 'MASTER_STOK', title: 'Master stok low/minus', note: 'Master stok', mode: 'stock_level' },
      { sheet: 'STOCK_SUBMISSION_QUEUE', title: 'Stock movement perlu review', note: 'Submission queue', needles: ['PENDING', 'NEEDS_REVIEW', 'NEEDS_APPROVAL'], exclude: ['CLOSED', 'DONE', 'APPROVED'] },
      { sheet: 'TASK_SUBMISSION', title: 'Task stok belum close', note: 'Task submission', needles: ['PENDING', 'OPEN', 'NEEDS'], exclude: ['CLOSED', 'DONE', 'APPROVED'] },
      { sheet: 'MATERIAL_USAGE_SUBMISSION', title: 'Pemakaian bahan pending', note: 'Material usage', needles: ['PENDING', 'OPEN', 'NEEDS'], exclude: ['CLOSED', 'DONE', 'APPROVED'] }
    ];
    const rows = [];
    let warn = 0;
    const staffTaskPack = kodReadStockDailyStaffTasks_(ss);
    if (staffTaskPack.checkedSheet) out.checkedSheets.push(staffTaskPack.checkedSheet);
    groups.forEach(function(g) {
      const sh = ss.getSheetByName(g.sheet);
      if (!sh) return;
      out.checkedSheets.push(g.sheet);
      const c = g.mode === 'stock_level' ? kodReadLowStockRows_(sh) : kodCountActionRows_(sh, g.needles, g.exclude || []);
      warn += c.count;
      if (c.count) rows.push({ source: 'Stock Log', type: g.sheet, title: g.title, count: c.count, note: g.note + (c.sampleStatus ? ' · contoh: ' + kodShort_(c.sampleStatus, 80) : ''), details: c.details || [] });
    });
    out.ok = true;
    out.warningCount = warn;
    out.rows = kodLimitRows_(kodSortActionRows_(rows), KOD_SAFE_LIMITS.maxDashboardRowsPerPanel);
    out.staffTasks = staffTaskPack.rows || [];
    out.staffTaskMeta = staffTaskPack.meta || {};
    out.warnings = kodLimitRows_(rows, 5).map(function(r) { return { source: 'Stock Log', title: r.title, note: r.count + ' item/pergerakan perlu cek', details: r.details || [] }; });
    out.note = out.checkedSheets.length ? 'Connected to stock levels + Stock Log daftar tugas.' : 'Connected, stock alert tabs not found.';
    return out;
  } catch (err) {
    return kodFail_(out, err);
  }
}

function kodBuildReceivingStockCountUrlV09_(receivingLineId) {
  const lineId = String(receivingLineId || '').trim();
  if (!lineId) return '';
  return KOD_ROUTE_LINKS.RECEIVING_WEBAPP + '?v90Action=ACTUAL_STOCK_COUNT&receivingLineId=' + encodeURIComponent(lineId);
}

function kodMapReceivingStockTaskRowsV09_(displayRows) {
  const out = { ok: false, rows: [], missingRequiredHeaders: [], headerMap: {}, openCount: 0, actionableCount: 0, closedCount: 0, skippedMissingLineId: 0 };
  const rows = Array.isArray(displayRows) ? displayRows : [];
  if (!rows.length) return out;
  const required = ['Task_ID','Task_Type','Status','Receiving_Line_ID','Receiving_Batch_ID','OrderBoard_ID','Item_Code','Item_Name','Purchase_Qty','Purchase_Unit','Stock_Unit','Evidence_Links_JSON'];
  const header = (rows[0] || []).map(function(v) { return String(v || '').trim(); });
  required.forEach(function(name) {
    const idx = header.indexOf(name);
    if (idx < 0) out.missingRequiredHeaders.push(name);
    else out.headerMap[name] = idx;
  });
  if (out.missingRequiredHeaders.length) return out;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const status = String(row[out.headerMap.Status] || '').trim().toUpperCase();
    if (status !== 'OPEN') {
      if (status) out.closedCount++;
      continue;
    }
    out.openCount++;
    const lineId = String(row[out.headerMap.Receiving_Line_ID] || '').trim();
    const mapped = {
      Task_ID: String(row[out.headerMap.Task_ID] || '').trim(),
      Task_Type: String(row[out.headerMap.Task_Type] || '').trim(),
      Status: 'OPEN',
      Receiving_Line_ID: lineId,
      Receiving_Batch_ID: String(row[out.headerMap.Receiving_Batch_ID] || '').trim(),
      OrderBoard_ID: String(row[out.headerMap.OrderBoard_ID] || '').trim(),
      Item_Code: String(row[out.headerMap.Item_Code] || '').trim(),
      Item_Name: String(row[out.headerMap.Item_Name] || '').trim(),
      Purchase_Qty: String(row[out.headerMap.Purchase_Qty] || '').trim(),
      Purchase_Unit: String(row[out.headerMap.Purchase_Unit] || '').trim(),
      Stock_Unit: String(row[out.headerMap.Stock_Unit] || '').trim(),
      Evidence_Links_JSON: String(row[out.headerMap.Evidence_Links_JSON] || '').trim(),
      actionable: !!lineId,
      actionLabel: lineId ? 'ISI QTY STOK AKTUAL' : '',
      actionUrl: lineId ? kodBuildReceivingStockCountUrlV09_(lineId) : ''
    };
    if (lineId) out.actionableCount++;
    else out.skippedMissingLineId++;
    out.rows.push(mapped);
  }
  out.ok = true;
  return out;
}

function kodReadReceivingStockCountTasks_() {
  const out = kodSourceOut_('RECEIVING');
  const sheetName = 'RECEIVING_STOCK_QTY_COUNT_TASK';
  try {
    const ss = SpreadsheetApp.openById(KOD_SOURCE_IDS.RECEIVING);
    const sh = ss.getSheetByName(sheetName);
    out.checkedSheets = [sheetName];
    if (!sh) {
      out.note = 'TASK_SHEET_MISSING';
      out.errorCode = 'TASK_SHEET_MISSING';
      out.meta = { sheetName: sheetName, requiredHeadersExact: true, readOnly: true };
      return out;
    }
    const lastRow = Math.min(sh.getLastRow(), KOD_SAFE_LIMITS.maxRowsPerSheet || 900);
    const lastCol = Math.min(sh.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet || 40);
    const display = (lastRow >= 1 && lastCol >= 1) ? sh.getRange(1, 1, lastRow, lastCol).getDisplayValues() : [];
    const mapped = kodMapReceivingStockTaskRowsV09_(display);
    out.ok = mapped.ok;
    out.rows = mapped.rows || [];
    out.pendingCount = mapped.actionableCount || 0;
    out.meta = {
      sheetName: sheetName,
      requiredHeadersExact: true,
      missingRequiredHeaders: mapped.missingRequiredHeaders || [],
      headerMap: mapped.headerMap || {},
      openCount: mapped.openCount || 0,
      actionableCount: mapped.actionableCount || 0,
      skippedMissingLineId: mapped.skippedMissingLineId || 0,
      readOnly: true,
      identityPolicy: 'EXACT_RECEIVING_LINE_ID_ONLY',
      routePolicy: 'RECEIVING_WEBAPP_V90_ACTUAL_STOCK_COUNT'
    };
    out.note = mapped.ok ? ('Connected. OPEN tasks: ' + (mapped.openCount || 0) + '; actionable exact line IDs: ' + (mapped.actionableCount || 0) + '.') : ('HEADER_MISMATCH: ' + (mapped.missingRequiredHeaders || []).join(', '));
    if (!mapped.ok) out.errorCode = 'HEADER_MISMATCH';
    return out;
  } catch (err) {
    return kodFail_(out, err);
  }
}

function kodReadInvoiceSummary_() {
  const out = kodSourceOut_('INVOICE_CUSTOMER');
  try {
    const ss = SpreadsheetApp.openById(KOD_SOURCE_IDS.INVOICE_CUSTOMER);
    const preferred = ['CUST_INV', 'INVOICE_LOG', 'RECEIVABLE_TRACKER', 'CUSTOMER_PAYMENT_LOG', 'SLIP_GAJI_V2'];
    const sheetNames = ss.getSheets().map(function(s) { return s.getName(); });
    let warn = 0;
    const rows = [];
    preferred.concat(sheetNames.slice(0, 6)).filter(function(name, idx, arr) { return arr.indexOf(name) === idx; }).forEach(function(name) {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      out.checkedSheets.push(name);
      const c = kodCountActionRows_(sh, ['UNPAID', 'PARTIAL', 'PENDING_CHECK', 'OVERDUE'], ['PAID_FULL', 'FULLY_PAID', 'CLOSED', 'CANCELLED', 'VOIDED']);
      if (c.count) {
        warn += c.count;
        rows.push({ source: 'Invoice', type: name, title: 'Invoice/piutang open', count: c.count, note: name + (c.sampleStatus ? ' · contoh: ' + kodShort_(c.sampleStatus, 70) : ''), details: c.details || [] });
      }
    });
    out.ok = true;
    out.warningCount = warn;
    out.rows = kodLimitRows_(kodSortActionRows_(rows), KOD_SAFE_LIMITS.maxDashboardRowsPerPanel);
    out.warnings = kodLimitRows_(rows, 3).map(function(r) { return { source: 'Invoice', title: r.title, note: r.count + ' row open di ' + r.type, details: r.details || [] }; });
    out.note = 'Connected. Invoice action groups only.';
    return out;
  } catch (err) {
    return kodFail_(out, err);
  }
}

function kodBuildOwnerBrief_(result) {
  const items = [];
  const receivingRows = result.receivingTasks || [];
  const receivingActionable = receivingRows.filter(function(r) { return r.actionable === true; });
  if (receivingActionable.length) {
    items.push({ tone: 'danger', title: receivingActionable.length + ' tugas Receiving perlu Cece kerjakan', note: 'Isi qty stok aktual lewat exact Receiving WebApp task.' });
  }

  const sourceAppPending = (result.pending || []).reduce(function(sum, r) { return sum + (Number(r.count) || 0); }, 0);
  if (sourceAppPending) items.push({ tone: 'warn', title: sourceAppPending + ' task menunggu di source app', note: 'Purchasing / Revenue / Cash tetap dikerjakan di WebApp sumber. Dashboard tidak menyalin mutation logic.' });

  const warnings = result.warnings || [];
  if (warnings.length) items.push({ tone: 'danger', title: warnings.length + ' peringatan penting', note: warnings[0].source + ' · ' + warnings[0].title });

  const prodRows = result.production || [];
  const realProd = prodRows.filter(function(r) { return r.pack !== 'ACTION'; });
  if (!realProd.length) items.push({ tone: 'warn', title: 'Produksi besok belum terbaca aman', note: 'Menu/pack besok belum terbaca sebagai data produksi aktif.' });

  if (!items.length) items.push({ tone: 'ok', title: 'Tidak ada tugas owner utama terbaca', note: 'Receiving actionable kosong dan tidak ada peringatan penting.' });
  return items.slice(0, 6);
}

function kodSourceOut_(source) {
  return { source: source, ok: false, note: '', checkedSheets: [], pendingCount: 0, warningCount: 0, refreshedAt: new Date().toISOString() };
}

function kodFail_(out, err) {
  out.ok = false;
  const msg = err && err.message ? err.message : String(err);
  if (/permission|authorization|auth|scope/i.test(msg)) {
    out.note = 'AUTH_GATE_PENDING';
    out.errorCode = 'AUTH_GATE_PENDING';
  } else {
    out.note = 'READ_FAILED: ' + kodShort_(msg, 90);
    out.errorCode = 'READ_FAILED';
  }
  return out;
}

function kodCountActionRows_(sheet, needles, excludes) {
  const lastRow = Math.min(sheet.getLastRow(), KOD_SAFE_LIMITS.statusScanMaxRows);
  const lastCol = Math.min(sheet.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet);
  if (lastRow < 2 || lastCol < 1) return { count: 0, sampleStatus: '', details: [] };
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headerRow = kodFindHeaderRow_(values, ['status', 'id', 'amount', 'staff', 'vendor', 'payment', 'task']);
  const hr = headerRow >= 0 ? headerRow : 0;
  const header = values[hr].map(function(h) { return String(h || '').trim().toLowerCase(); });
  const statusCols = [];
  const sampleCols = [];
  header.forEach(function(h, idx) {
    const hn = String(h || '').trim().toLowerCase();
    // Do not treat Transaction_Date / Action_Label / Action_Target as status columns.
    // Only actual status/state columns may decide pending state.
    if (/^(status|state|approval_status|payment_status|verify_status|posted_status|task_status|queue_status|sync_status|settlement_status|finance_sync_status|clean_status|final_status|klikbca_status|bca_flow)$/i.test(hn)) statusCols.push(idx);
    if (/id|status|vendor|staff|pic|amount|jumlah|payment|proof|note|reason|source|customer|invoice/i.test(hn)) sampleCols.push(idx);
  });
  if (!statusCols.length) statusCols.push(0);
  const upperNeedles = (needles || []).map(function(n) { return String(n).toUpperCase(); });
  const upperExcludes = (excludes || []).concat(KOD_FINAL_STATUSES || []).map(function(n) { return String(n).toUpperCase(); });
  let count = 0;
  let sampleStatus = '';
  const details = [];
  for (let r = hr + 1; r < values.length; r++) {
    const statusText = statusCols.map(function(c) { return String(values[r][c] || '').trim(); }).filter(String).join(' / ');
    const rowText = (sampleCols.length ? sampleCols : statusCols).map(function(c) { return String(values[r][c] || '').trim(); }).filter(String).join(' / ');
    const statusUpper = statusText.toUpperCase();
    const candidateUpper = (statusText || rowText).toUpperCase();
    const excluded = upperExcludes.some(function(n) { return n && kodTokenHit_(statusUpper, n); });
    if (excluded) continue;
    const hit = upperNeedles.some(function(n) { return n && kodTokenHit_(candidateUpper, n); });
    if (hit) {
      count++;
      if (!sampleStatus) sampleStatus = kodShort_(rowText || statusText, 120);
      if (details.length < 10) details.push({ sheet: sheet.getName(), row: r + 1, status: statusText, summary: kodShort_(rowText || statusText, 180) });
    }
  }
  return { count: count, sampleStatus: sampleStatus, details: details };
}

function kodBuildStaffTasksFromQueue_(sheet, needles, excludes) {
  const lastRow = Math.min(sheet.getLastRow(), KOD_SAFE_LIMITS.statusScanMaxRows);
  const lastCol = Math.min(sheet.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet);
  if (lastRow < 2 || lastCol < 1) return [];
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headerRow = kodFindHeaderRow_(values, ['staff', 'pic', 'assignee', 'status', 'task']);
  const hr = headerRow >= 0 ? headerRow : 0;
  const header = values[hr].map(function(h) { return String(h || '').trim().toLowerCase(); });
  const idxStaff = kodFirstHeader_(header, ['staff', 'pic', 'assignee', 'owner', 'handler']);
  const idxStatus = kodFirstHeader_(header, ['status', 'task_status', 'state']);
  const idxTask = kodFirstHeader_(header, ['task', 'action', 'reason', 'note', 'type', 'source']);
  const result = [];
  const countByStaff = {};
  const sampleByStaff = {};
  for (let r = hr + 1; r < values.length; r++) {
    const status = idxStatus >= 0 ? values[r][idxStatus] : values[r].join(' / ');
    const statusUpper = String(status || '').toUpperCase();
    const excluded = (excludes || []).concat(KOD_FINAL_STATUSES || []).some(function(n) { return kodTokenHit_(statusUpper, n); });
    if (excluded) continue;
    const hit = (needles || []).some(function(n) { return kodTokenHit_(statusUpper, n); });
    if (!hit) continue;
    const staff = (idxStaff >= 0 ? String(values[r][idxStaff] || '').trim() : '') || 'Staff fix';
    const task = (idxTask >= 0 ? String(values[r][idxTask] || '').trim() : '') || status;
    countByStaff[staff] = (countByStaff[staff] || 0) + 1;
    if (!sampleByStaff[staff]) sampleByStaff[staff] = task;
  }
  Object.keys(countByStaff).forEach(function(staff) {
    result.push({ staff: staff, task: countByStaff[staff] + ' task', note: kodShort_(sampleByStaff[staff] || 'Perlu action', 90), count: countByStaff[staff] });
  });
  return result;
}

function kodMergeStaffTasks_(tasks) {
  const map = {};
  (tasks || []).forEach(function(t) {
    const staff = t.staff || 'Staff';
    if (!map[staff]) map[staff] = { staff: staff, count: 0, samples: [] };
    map[staff].count += t.count || kodAsNumber_(t.task) || 1;
    if (t.note) map[staff].samples.push(t.note);
  });
  return Object.keys(map).map(function(k) {
    const m = map[k];
    return { staff: m.staff, task: m.count + ' task', note: kodUnique_(m.samples).slice(0, 2).join(' · '), count: m.count };
  }).sort(function(a, b) { return b.count - a.count; });
}


function kodReadStockDailyStaffTasks_(ss) {
  const sh = ss.getSheetByName('TASK_SUBMISSION');
  const out = { rows: [], checkedSheet: '', meta: { mode: 'NO_TASK_SUBMISSION', targetDate: '', activeCount: 0 } };
  if (!sh) return out;
  out.checkedSheet = 'TASK_SUBMISSION';
  const lastRow = Math.min(sh.getLastRow(), KOD_SAFE_LIMITS.statusScanMaxRows);
  const lastCol = Math.min(sh.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet);
  if (lastRow < 2 || lastCol < 10) return out;

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headerRow = kodFindHeaderRow_(display, ['TASK_ID', 'TANGGAL', 'TASK_TEXT', 'PIC_PENGERJA']);
  const hr = headerRow >= 0 ? headerRow : 0;
  const header = display[hr].map(function(h) { return String(h || '').trim().toLowerCase(); });
  const idxDate = kodFirstHeader_(header, ['tanggal', 'task_date', 'date', 'tanggal_catering']);
  const idxPic = kodFirstHeader_(header, ['pic_pengerja', 'pic', 'staff', 'assignee', 'assigned_to']);
  const idxTask = kodFirstHeader_(header, ['task_text', 'task', 'tugas', 'action']);
  const idxProgram = kodFirstHeader_(header, ['program', 'section_title', 'section_key']);
  const idxMenu = kodFirstHeader_(header, ['menu_name', 'menu']);
  const idxQty = kodFirstHeader_(header, ['target_qty', 'qty', 'total_produksi']);
  const idxDone = kodFirstHeader_(header, ['done', 'is_done', 'completed']);
  const idxStatus = kodFirstHeader_(header, ['status', 'task_status', 'state']);
  const idxId = kodFirstHeader_(header, ['task_id', 'row_id', 'submission_id']);

  if (idxTask < 0) return out;

  const candidates = [];
  for (let r = hr + 1; r < values.length; r++) {
    const rawTask = String(display[r][idxTask] || '').trim();
    if (!rawTask || kodLooksLikeNoise_(rawTask)) continue;
    const status = idxStatus >= 0 ? String(display[r][idxStatus] || '').trim() : '';
    const doneRaw = idxDone >= 0 ? values[r][idxDone] : false;
    const doneText = idxDone >= 0 ? String(display[r][idxDone] || '').trim().toUpperCase() : '';
    if (kodIsTaskDoneOrClosed_(status, doneRaw, doneText)) continue;
    const dateKey = idxDate >= 0 ? (kodDateKey_(values[r][idxDate]) || kodDateKey_(display[r][idxDate])) : '';
    const pic = idxPic >= 0 ? String(display[r][idxPic] || '').trim() : '';
    const program = idxProgram >= 0 ? String(display[r][idxProgram] || '').trim() : '';
    const menu = idxMenu >= 0 ? String(display[r][idxMenu] || '').trim() : '';
    const qty = idxQty >= 0 ? String(display[r][idxQty] || '').trim() : '';
    candidates.push({
      dateKey: dateKey,
      pic: pic || 'Belum assign',
      taskText: rawTask,
      program: program,
      menu: menu,
      qty: qty,
      status: status || 'OPEN',
      sheet: sh.getName(),
      row: r + 1,
      id: idxId >= 0 ? String(display[r][idxId] || '').trim() : ''
    });
  }

  out.meta.activeCount = candidates.length;
  if (!candidates.length) {
    out.meta.mode = 'NO_ACTIVE_TASK';
    return out;
  }

  const tomorrowKey = kodDateOffsetKey_(1);
  const todayKey = kodDateOffsetKey_(0);
  const byDate = {};
  candidates.forEach(function(t) {
    const k = t.dateKey || 'NO_DATE';
    if (!byDate[k]) byDate[k] = [];
    byDate[k].push(t);
  });
  let targetKey = tomorrowKey;
  let chosen = byDate[tomorrowKey] || [];
  let mode = 'TOMORROW_TASKS';
  if (!chosen.length && byDate[todayKey] && byDate[todayKey].length) {
    targetKey = todayKey;
    chosen = byDate[todayKey];
    mode = 'TODAY_TASKS';
  }
  if (!chosen.length) {
    const dateKeys = Object.keys(byDate).filter(function(k) { return /^20\d{2}-\d{2}-\d{2}$/.test(k); }).sort();
    targetKey = dateKeys.length ? dateKeys[dateKeys.length - 1] : 'NO_DATE';
    chosen = byDate[targetKey] || candidates;
    mode = 'LATEST_ACTIVE_TASKS';
  }

  out.meta.mode = mode;
  out.meta.targetDate = targetKey;
  out.rows = kodGroupStockTasksByPic_(chosen, targetKey, mode);
  return out;
}

function kodIsTaskDoneOrClosed_(status, doneRaw, doneText) {
  const st = String(status || '').toUpperCase();
  const done = doneRaw === true || doneText === 'TRUE' || doneText === 'YES' || doneText === 'DONE' || doneText === 'SELESAI';
  if (done) return true;
  if (/DONE|CLOSED|APPROVED|REVIEWED|COMPLETED|CANCELLED|VOIDED/.test(st)) return true;
  if (/EXPIRED_NO_REVIEW/.test(st)) return true;
  return false;
}

function kodGroupStockTasksByPic_(tasks, targetKey, mode) {
  const map = {};
  (tasks || []).forEach(function(t) {
    const pic = String(t.pic || 'Belum assign').trim() || 'Belum assign';
    if (!map[pic]) map[pic] = { staff: pic, count: 0, samples: [], details: [] };
    map[pic].count++;
    if (t.taskText && map[pic].samples.length < 3) map[pic].samples.push(t.taskText);
    if (map[pic].details.length < 20) {
      map[pic].details.push({
        staff: pic,
        task: t.taskText,
        program: t.program,
        menu: kodShort_(String(t.menu || '').replace(/\n+/g, ' / '), 140),
        qty: t.qty,
        tanggal: targetKey,
        status: t.status,
        sheet: t.sheet,
        row: t.row,
        id: t.id
      });
    }
  });
  return Object.keys(map).map(function(pic) {
    const m = map[pic];
    const labelDate = targetKey && targetKey !== 'NO_DATE' ? targetKey : 'tanggal belum ada';
    const sourceNote = mode === 'LATEST_ACTIVE_TASKS' ? 'Daftar tugas aktif terakhir' : 'Daftar tugas harian';
    return {
      staff: m.staff,
      task: m.count + ' tugas',
      count: m.count,
      note: sourceNote + ' ' + labelDate + ' · ' + kodUnique_(m.samples).slice(0, 2).join(' · '),
      details: m.details
    };
  }).sort(function(a, b) { return b.count - a.count || String(a.staff).localeCompare(String(b.staff)); }).slice(0, KOD_SAFE_LIMITS.maxDashboardRowsPerPanel);
}

function kodReadLowStockRows_(sheet) {
  const lastRow = Math.min(sheet.getLastRow(), KOD_SAFE_LIMITS.statusScanMaxRows);
  const lastCol = Math.min(sheet.getLastColumn(), KOD_SAFE_LIMITS.maxColsPerSheet);
  if (lastRow < 2 || lastCol < 1) return { count: 0, sampleStatus: '', details: [] };
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headerRow = kodFindHeaderRow_(display, ['bahan', 'item', 'stok', 'stock', 'min', 'status']);
  const hr = headerRow >= 0 ? headerRow : 0;
  const header = display[hr].map(function(h) { return String(h || '').trim().toLowerCase(); });
  const idxItem = kodFirstHeader_(header, ['bahan', 'item', 'nama', 'product', 'material']);
  const idxStock = kodFirstHeader_(header, ['stok akhir', 'stock akhir', 'current stock', 'stok', 'stock', 'qty']);
  const idxMin = kodFirstHeader_(header, ['minimum', 'min stock', 'min', 'par']);
  const idxStatus = kodFirstHeader_(header, ['status', 'warning', 'alert']);
  let count = 0;
  let sample = '';
  const details = [];
  for (let r = hr + 1; r < values.length; r++) {
    const item = idxItem >= 0 ? String(display[r][idxItem] || '').trim() : '';
    if (!item || kodLooksLikeNoise_(item)) continue;
    const status = idxStatus >= 0 ? String(display[r][idxStatus] || '').toUpperCase() : '';
    const stock = idxStock >= 0 ? kodAsNumber_(values[r][idxStock] || display[r][idxStock]) : 0;
    const min = idxMin >= 0 ? kodAsNumber_(values[r][idxMin] || display[r][idxMin]) : 0;
    const low = /LOW|MINUS|CRITICAL|KURANG|HABIS|WARNING/.test(status) || (idxStock >= 0 && idxMin >= 0 && stock <= min && min > 0) || (idxStock >= 0 && stock < 0);
    if (!low) continue;
    count++;
    const line = item + (idxStock >= 0 ? ' sisa ' + display[r][idxStock] : '') + (idxMin >= 0 ? ' min ' + display[r][idxMin] : '');
    if (!sample) sample = line;
    if (details.length < 10) details.push({ sheet: sheet.getName(), row: r + 1, item: item, stock: idxStock >= 0 ? display[r][idxStock] : '', min: idxMin >= 0 ? display[r][idxMin] : '', status: status || (stock <= min ? 'LOW' : '') });
  }
  return { count: count, sampleStatus: sample, details: details };
}

function kodReadComputedCashSnapshot_(ss) {
  const accountRows = kodReadAccountMasterRows_(ss);
  const byId = {};
  const aliases = {};
  accountRows.forEach(function(a) {
    byId[a.id] = {
      id: a.id,
      name: a.name || a.id,
      type: a.type || '',
      opening: a.opening,
      hasOpening: a.hasOpening,
      ledgerDelta: 0,
      cashboxDelta: 0,
      pendingOut: 0,
      ledgerRows: 0,
      cashboxRows: 0
    };
    kodRegisterAccountAlias_(aliases, a.id, a.id);
    kodRegisterAccountAlias_(aliases, String(a.id).replace(/_/g, ' '), a.id);
    kodRegisterAccountAlias_(aliases, a.name, a.id);
  });
  kodRegisterDefaultAccountAliases_(aliases);

  const checkedSheets = [];
  const ledger = ss.getSheetByName('CASH_BANK_LEDGER');
  if (ledger) {
    checkedSheets.push('CASH_BANK_LEDGER');
    kodApplyCashBankLedgerDeltas_(ledger, byId, aliases);
  }
  const cashbox = ss.getSheetByName('CASHBOX_MOVEMENT');
  if (cashbox) {
    checkedSheets.push('CASHBOX_MOVEMENT');
    kodApplyCashboxMovementDeltas_(cashbox, byId, aliases);
  }
  const request = ss.getSheetByName('PAYMENT_REQUEST_QUEUE');
  if (request) {
    checkedSheets.push('PAYMENT_REQUEST_QUEUE');
    kodApplyPaymentRequestPending_(request, byId, aliases);
  }

  const preferredOrder = ['BCA_KALMA', 'MANDIRI_CECE_QRIS', 'PETTYCASH_KALMA', 'PETTYCASH_CANTEEN', 'CECE_PRIBADI', 'CECE_TALANGAN', 'OVO_KALMA', 'QRIS_PENDING', 'CASH_HELD_BY_STAFF', 'SHOPEE_CECE', 'SAMPOERNA_SETTLEMENT_PENDING'];
  const accounts = preferredOrder.map(function(id) { return byId[id]; }).filter(Boolean).map(function(a) {
    const finalBalance = kodRoundMoney_((a.opening || 0) + (a.ledgerDelta || 0) + (a.cashboxDelta || 0));
    const availableAfterPending = kodRoundMoney_(finalBalance - (a.pendingOut || 0));
    const componentNote = 'Opening ' + kodFormatRp_(a.opening || 0) +
      ' + Ledger ' + kodFormatSignedRp_(a.ledgerDelta || 0) +
      ' + Cashbox ' + kodFormatSignedRp_(a.cashboxDelta || 0) +
      ((a.pendingOut || 0) ? ' · Pending bayar ' + kodFormatRp_(a.pendingOut || 0) : '');
    let warning = '';
    if (!a.hasOpening && Math.abs(finalBalance) > 0) warning = 'Opening kosong tapi ada movement; saldo belum final.';
    if (finalBalance < 0) warning = 'Saldo hitung negatif; cek bank statement / duplicate posting.';
    if (availableAfterPending < 0 && finalBalance >= 0) warning = 'Saldo akan minus jika pending payment dibayar.';
    if (a.id === 'CECE_PRIBADI' && !a.hasOpening && finalBalance === 0) warning = '';
    return {
      account: a.id,
      accountName: a.name,
      type: a.type,
      balanceValue: finalBalance,
      balance: kodFormatRp_(finalBalance),
      availableAfterPendingValue: availableAfterPending,
      availableAfterPending: kodFormatRp_(availableAfterPending),
      openingValue: a.opening || 0,
      ledgerDelta: a.ledgerDelta || 0,
      cashboxDelta: a.cashboxDelta || 0,
      pendingOut: a.pendingOut || 0,
      ledgerRows: a.ledgerRows || 0,
      cashboxRows: a.cashboxRows || 0,
      note: componentNote,
      warning: warning,
      details: [{
        account: a.id,
        current: kodFormatRp_(finalBalance),
        availableAfterPending: kodFormatRp_(availableAfterPending),
        opening: kodFormatRp_(a.opening || 0),
        ledgerDelta: kodFormatSignedRp_(a.ledgerDelta || 0),
        cashboxDelta: kodFormatSignedRp_(a.cashboxDelta || 0),
        pendingPayment: kodFormatRp_(a.pendingOut || 0),
        ledgerRows: a.ledgerRows || 0,
        cashboxRows: a.cashboxRows || 0,
        warning: warning || '-'
      }]
    };
  }).filter(function(a) { return !(a.account === 'CECE_PRIBADI' && !a.warning && a.balanceValue === 0); });
  return { accounts: accounts, checkedSheets: checkedSheets, meta: { formula: 'opening_balance + full_cash_bank_ledger_delta + cashbox_movement_delta; pending shown separately', accountCount: accounts.length } };
}

function kodReadAccountMasterRows_(ss) {
  const sh = ss.getSheetByName('ACCOUNT_MASTER');
  if (!sh) return [];
  const lastRow = Math.min(sh.getLastRow(), 160);
  const lastCol = Math.min(sh.getLastColumn(), 24);
  if (lastRow < 2) return [];
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headerRow = kodFindHeaderRow_(display, ['account', 'opening', 'balance']);
  const hr = headerRow >= 0 ? headerRow : 0;
  const header = display[hr].map(function(h) { return String(h || '').trim().toLowerCase(); });
  const idxId = kodFirstHeader_(header, ['account id', 'account_id', 'id', 'account']);
  const idxName = kodFirstHeader_(header, ['account name', 'account_name', 'rekening', 'cashbox', 'wallet', 'nama']);
  const idxType = kodFirstHeader_(header, ['account type', 'account_type', 'type']);
  const idxOpening = kodFirstHeader_(header, ['opening balance', 'opening_balance', 'opening']);
  const result = [];
  for (let i = hr + 1; i < values.length; i++) {
    const id = String(display[i][idxId >= 0 ? idxId : 0] || '').trim();
    if (!id) continue;
    const rawOpening = idxOpening >= 0 ? values[i][idxOpening] : '';
    const displayOpening = idxOpening >= 0 ? display[i][idxOpening] : '';
    result.push({ id: id, name: String(display[i][idxName >= 0 ? idxName : 1] || id).trim(), type: String(display[i][idxType >= 0 ? idxType : 2] || '').trim(), opening: kodAsNumber_(rawOpening || displayOpening), hasOpening: String(displayOpening || '').trim() !== '' });
  }
  return result;
}

function kodApplyCashBankLedgerDeltas_(sheet, byId, aliases) {
  const lastRow = Math.min(sheet.getLastRow(), KOD_SAFE_LIMITS.cashLedgerMaxRows || 5000);
  const lastCol = Math.min(sheet.getLastColumn(), (KOD_SAFE_LIMITS.maxColsPerSheet || 40) + 20);
  if (lastRow < 2 || lastCol < 1) return;
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headerRow = kodFindHeaderRow_(display, ['txn', 'amount', 'account', 'status']);
  const hr = headerRow >= 0 ? headerRow : 0;
  const header = display[hr].map(function(h) { return String(h || '').trim().toLowerCase(); });
  const idxFrom = kodFirstHeader_(header, ['from account', 'from_account', 'source account', 'from']);
  const idxTo = kodFirstHeader_(header, ['to account', 'to_account', 'target account', 'destination', 'to']);
  const idxDirection = kodFirstHeader_(header, ['direction']);
  const idxTxnType = kodFirstHeader_(header, ['txn type', 'txn_type', 'transaction type', 'transaction_type', 'type']);
  const idxAmount = kodFirstHeader_(header, ['amount', 'jumlah', 'nominal']);
  const idxStatus = kodFirstHeader_(header, ['status']);
  const idxAccount = kodFirstHeader_(header, ['account']);
  if (idxAmount < 0) return;

  for (let r = hr + 1; r < values.length; r++) {
    const status = idxStatus >= 0 ? display[r][idxStatus] : '';
    if (!kodIsCashSnapshotPostedStatus_(status)) continue;
    const signedAmount = kodAsNumber_(values[r][idxAmount] !== '' && values[r][idxAmount] !== null ? values[r][idxAmount] : display[r][idxAmount]);
    if (!signedAmount) continue;
    const amount = Math.abs(signedAmount);
    const fromId = kodCanonicalAccount_(idxFrom >= 0 ? display[r][idxFrom] : '', aliases);
    const toId = kodCanonicalAccount_(idxTo >= 0 ? display[r][idxTo] : '', aliases);
    const accountId = kodCanonicalAccount_(idxAccount >= 0 ? display[r][idxAccount] : '', aliases);
    const direction = String(idxDirection >= 0 ? display[r][idxDirection] : '').toUpperCase();
    const txnType = String(idxTxnType >= 0 ? display[r][idxTxnType] : '').toUpperCase();
    const rowText = (direction + ' ' + txnType + ' ' + String(status || '').toUpperCase()).toUpperCase();

    if (/REVERSAL/.test(rowText) && fromId && toId) {
      // Reversal rows are already explicit from/to rows. Use signed direction conservatively.
      kodAddLedgerDelta_(byId, fromId, signedAmount < 0 ? amount : -amount);
      kodAddLedgerDelta_(byId, toId, signedAmount < 0 ? -amount : amount);
      continue;
    }

    if (/OUT|MONEY_OUT|PAYMENT|EXPENSE|SUPPLIER_PAYMENT|PAYROLL|WITHDRAW/.test(rowText)) {
      if (fromId) kodAddLedgerDelta_(byId, fromId, -amount);
      else if (accountId) kodAddLedgerDelta_(byId, accountId, -amount);
      continue;
    }
    if (/IN|MONEY_IN|DEPOSIT|REVENUE|SALES|SETTLEMENT|TOPUP/.test(rowText)) {
      if (toId) kodAddLedgerDelta_(byId, toId, amount);
      else if (accountId) kodAddLedgerDelta_(byId, accountId, amount);
      continue;
    }
    if (/TRANSFER|MOVE/.test(rowText)) {
      if (fromId) kodAddLedgerDelta_(byId, fromId, -amount);
      if (toId) kodAddLedgerDelta_(byId, toId, amount);
      continue;
    }
    if (fromId && toId) {
      kodAddLedgerDelta_(byId, fromId, -amount);
      kodAddLedgerDelta_(byId, toId, amount);
    } else if (fromId) {
      kodAddLedgerDelta_(byId, fromId, -amount);
    } else if (toId) {
      kodAddLedgerDelta_(byId, toId, amount);
    } else if (accountId) {
      kodAddLedgerDelta_(byId, accountId, signedAmount);
    }
  }
}

function kodApplyCashboxMovementDeltas_(sheet, byId, aliases) {
  const lastRow = Math.min(sheet.getLastRow(), KOD_SAFE_LIMITS.cashboxMovementMaxRows || 3000);
  const lastCol = Math.min(sheet.getLastColumn(), (KOD_SAFE_LIMITS.maxColsPerSheet || 40) + 8);
  if (lastRow < 2 || lastCol < 1) return;
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headerRow = kodFindHeaderRow_(display, ['amount', 'cashbox', 'status']);
  const hr = headerRow >= 0 ? headerRow : 0;
  const header = display[hr].map(function(h) { return String(h || '').trim().toLowerCase(); });
  const idxFrom = kodFirstHeader_(header, ['from cashbox', 'from_cashbox', 'from source', 'from_source', 'source']);
  const idxTo = kodFirstHeader_(header, ['to cashbox', 'to_cashbox', 'to destination', 'to_destination', 'destination']);
  const idxAmount = kodFirstHeader_(header, ['amount', 'jumlah', 'nominal']);
  const idxStatus = kodFirstHeader_(header, ['status']);
  if (idxAmount < 0) return;
  for (let r = hr + 1; r < values.length; r++) {
    const status = idxStatus >= 0 ? display[r][idxStatus] : '';
    if (!kodIsCashSnapshotPostedStatus_(status)) continue;
    const amount = Math.abs(kodAsNumber_(values[r][idxAmount] !== '' && values[r][idxAmount] !== null ? values[r][idxAmount] : display[r][idxAmount]));
    if (!amount) continue;
    const fromId = kodCanonicalAccount_(idxFrom >= 0 ? display[r][idxFrom] : '', aliases);
    const toId = kodCanonicalAccount_(idxTo >= 0 ? display[r][idxTo] : '', aliases);
    kodAddCashboxDelta_(byId, fromId, -amount);
    kodAddCashboxDelta_(byId, toId, amount);
  }
}

function kodApplyPaymentRequestPending_(sheet, byId, aliases) {
  const lastRow = Math.min(sheet.getLastRow(), KOD_SAFE_LIMITS.statusScanMaxRows || 1400);
  const lastCol = Math.min(sheet.getLastColumn(), (KOD_SAFE_LIMITS.maxColsPerSheet || 40) + 8);
  if (lastRow < 2 || lastCol < 1) return;
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headerRow = kodFindHeaderRow_(display, ['amount', 'status', 'account']);
  const hr = headerRow >= 0 ? headerRow : 0;
  const header = display[hr].map(function(h) { return String(h || '').trim().toLowerCase(); });
  const idxAccount = kodFirstHeader_(header, ['requested from account', 'requested_from_account', 'from account', 'account']);
  const idxAmount = kodFirstHeader_(header, ['amount', 'jumlah', 'nominal']);
  const idxStatus = kodFirstHeader_(header, ['status']);
  if (idxAmount < 0 || idxAccount < 0) return;
  for (let r = hr + 1; r < values.length; r++) {
    const status = String(idxStatus >= 0 ? display[r][idxStatus] : '').toUpperCase();
    if (!/NEEDS_PAYMENT|READY_TO_PAY|WAITING_PAYMENT|SUBMITTED/.test(status)) continue;
    if (/PAID|POSTED|VOID|CANCEL|REJECT|CLOSED/.test(status)) continue;
    const id = kodCanonicalAccount_(display[r][idxAccount], aliases);
    const amount = Math.abs(kodAsNumber_(values[r][idxAmount] !== '' && values[r][idxAmount] !== null ? values[r][idxAmount] : display[r][idxAmount]));
    if (id && byId[id] && amount) byId[id].pendingOut += amount;
  }
}

function kodAddLedgerDelta_(byId, id, amount) { if (id && byId[id]) { byId[id].ledgerDelta += amount || 0; byId[id].ledgerRows = (byId[id].ledgerRows || 0) + (amount ? 1 : 0); } }
function kodAddCashboxDelta_(byId, id, amount) { if (id && byId[id]) { byId[id].cashboxDelta += amount || 0; byId[id].cashboxRows = (byId[id].cashboxRows || 0) + (amount ? 1 : 0); } }
function kodCanonicalAccount_(value, aliases) { const raw = String(value || '').trim(); if (!raw) return ''; const key = raw.toUpperCase().replace(/[_\-]+/g, ' ').replace(/\s+/g, ' ').trim(); return aliases[key] || aliases[raw.toUpperCase()] || raw; }
function kodRegisterAccountAlias_(aliases, alias, id) { alias = String(alias || '').trim(); if (!alias || !id) return; aliases[alias.toUpperCase()] = id; aliases[alias.toUpperCase().replace(/[_\-]+/g, ' ').replace(/\s+/g, ' ').trim()] = id; }
function kodRegisterDefaultAccountAliases_(aliases) { ['BCA Kalma|BCA_KALMA','Petty Cash Kalma|PETTYCASH_KALMA','Petty Cash Canteen|PETTYCASH_CANTEEN','Mandiri Cece QRIS|MANDIRI_CECE_QRIS','OVO Kalma|OVO_KALMA','QRIS Pending|QRIS_PENDING','Cece Pribadi|CECE_PRIBADI','Cece Talangan|CECE_TALANGAN','Shopee Cece|SHOPEE_CECE','Sampoerna Settlement Pending|SAMPOERNA_SETTLEMENT_PENDING'].forEach(function(x){ const p=x.split('|'); kodRegisterAccountAlias_(aliases,p[0],p[1]); }); }
function kodIsCashSnapshotPostedStatus_(status) { const s = String(status || '').toUpperCase().trim(); if (!s) return false; if (/VOID|CANCEL|REJECT|DRAFT|TEST/.test(s)) return false; return /POSTED|PAID|DEPOSITED|SETTLED|PENDING_SETTLEMENT|RECORD_ONLY|SUBMITTED|APPROVED/.test(s); }

function kodFindHeaderRow_(display, keys) {
  const max = Math.min(12, display.length);
  for (let r = 0; r < max; r++) {
    const row = (display[r] || []).map(function(x) { return String(x || '').toLowerCase(); });
    let hits = 0;
    (keys || []).forEach(function(k) { const nk = String(k).toLowerCase().replace(/[ _\-]/g, ''); if (row.some(function(c) { return c.replace(/[ _\-]/g, '').indexOf(nk) >= 0; })) hits++; });
    if (hits >= 2) return r;
  }
  return -1;
}
function kodFirstHeader_(header, keys) { for (let i = 0; i < header.length; i++) { const h = String(header[i] || '').toLowerCase().replace(/[ _\-]/g, ''); for (let k = 0; k < keys.length; k++) { const key = String(keys[k] || '').toLowerCase().replace(/[ _\-]/g, ''); if (h.indexOf(key) >= 0) return i; } } return -1; }
function kodSortActionRows_(rows) { return (rows || []).sort(function(a, b) { return (Number(b.count) || 0) - (Number(a.count) || 0); }); }
function kodTokenHit_(text, token) { text = String(text || '').toUpperCase(); token = String(token || '').toUpperCase(); if (!token) return false; const escaped = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); const re = new RegExp('(^|[^A-Z0-9])' + escaped + '([^A-Z0-9]|$)'); return re.test(text); }
function kodDateOffsetKey_(offsetDays) { const now = new Date(); const tz = Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Asia/Jakarta'; const base = new Date(now.getTime() + (offsetDays || 0) * 24 * 60 * 60 * 1000); return Utilities.formatDate(base, tz, 'yyyy-MM-dd'); }
function kodDateKey_(value) { if (value === null || value === undefined || value === '') return ''; const tz = Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Asia/Jakarta'; if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return Utilities.formatDate(value, tz, 'yyyy-MM-dd'); if (typeof value === 'number') { const d = new Date(Date.UTC(1899, 11, 30)); d.setUTCDate(d.getUTCDate() + Math.floor(value)); return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); } const s = String(value).trim(); if (/^\d+(\.0+)?$/.test(s) && Number(s) > 30000) return kodDateKey_(Number(s)); const iso = s.match(/(20\d{2})[-\/](\d{1,2})[-\/](\d{1,2})/); if (iso) return iso[1] + '-' + kodPad2_(iso[2]) + '-' + kodPad2_(iso[3]); const dmy = s.match(/(\d{1,2})[-\/](\d{1,2})[-\/](20\d{2})/); if (dmy) return dmy[3] + '-' + kodPad2_(dmy[2]) + '-' + kodPad2_(dmy[1]); const parsed = new Date(s); if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd'); return ''; }
function kodRowContainsDateKey_(row, tomorrowKey) { for (let i = 0; i < row.length; i++) { if (kodDateKey_(row[i]) === tomorrowKey) return tomorrowKey; } return ''; }
function kodPad2_(x) { return String(x).padStart(2, '0'); }
function kodAsNumber_(v) { if (typeof v === 'number' && isFinite(v)) return v; const raw = String(v || '').trim(); if (/^\d{1,2}([./-])\d{1,2}\d{2,4}$/.test(raw)) return 0; const s = raw.replace(/[^0-9,.-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'); const n = Number(s); return isFinite(n) ? n : 0; }
function kodFormatNumber_(n) { n = kodAsNumber_(n); if (!n) return '0'; return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function kodRoundMoney_(n) { n = kodAsNumber_(n); return Math.round(n); }
function kodFormatRp_(n) { n = kodRoundMoney_(n); const sign = n < 0 ? '-' : ''; return sign + 'Rp ' + String(Math.abs(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function kodFormatSignedRp_(n) { n = kodRoundMoney_(n); const sign = n > 0 ? '+' : (n < 0 ? '-' : ''); return sign + 'Rp ' + String(Math.abs(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function kodLimitRows_(rows, limit) { rows = rows || []; limit = limit || 6; if (rows.length <= limit) return rows; const sliced = rows.slice(0, limit); sliced.push({ source: 'Dashboard', title: 'Ada ' + (rows.length - limit) + ' kategori lain', count: '+', note: 'Buka detail panel di dashboard atau modul asal jika perlu action.', details: rows.slice(limit, limit + 8) }); return sliced; }
function kodShort_(s, maxLen) { s = String(s || '').replace(/\s+/g, ' ').trim(); maxLen = maxLen || 120; return s.length > maxLen ? s.slice(0, maxLen - 1) + '…' : s; }
function kodUnique_(arr) { const seen = {}; return (arr || []).filter(function(x) { x = String(x || '').trim(); if (!x || seen[x]) return false; seen[x] = true; return true; }); }
function kodNormalizeKey_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim(); }
function kodLooksLikeNoise_(s) { s = String(s || '').trim(); if (!s) return true; if (/^(total|grand total|jumlah|date|tanggal|menu|program|customer|nama|status)$/i.test(s)) return true; if (s.length > 180) return true; return false; }
function kodProgramFromSheet_(name) { name = String(name || '').toUpperCase(); if (name.indexOf('DAILY') >= 0 || /^D_/.test(name)) return 'Daily Catering'; if (name.indexOf('HLT') >= 0 || /^H_/.test(name)) return 'Healthy Catering'; if (name.indexOf('KIDS') >= 0 || /^K_/.test(name) || name.indexOf('BIGSTEPS') >= 0) return 'Kids Catering'; if (name.indexOf('TEEN') >= 0 || /^T_/.test(name)) return 'Teen Catering'; if (name.indexOf('SAMPOERNA') >= 0 || name.indexOf('CANTEEN') >= 0) return 'Canteen'; return name; }
function kodIsCancelledOrFinalProduction_(status) { const s = String(status || '').toUpperCase(); return /CANCEL|VOID|LIBUR|PAUSE|SKIP|DONE|CLOSED|TEST/.test(s); }
function kodGuessMenuFromRow_(row) { const txt = (row || []).map(function(x) { return String(x || '').trim(); }).filter(String).filter(function(x) { return !kodLooksLikeNoise_(x) && !/^\d+[\d.,]*$/.test(x); }).join(' · '); return kodShort_(txt, 90); }
function kodGuessQuantityFromRow_(values, display) {
  for (let i = 0; i < Math.min(values.length, 18); i++) {
    const n = kodSafeProductionPack_(values[i], display[i], 'row-scan');
    if (n > 0) return n;
  }
  return 0;
}
function kodSafeProductionPack_(value, displayValue, context) {
  if (value === null || value === undefined || value === '') return 0;
  if (Object.prototype.toString.call(value) === '[object Date]') return 0;
  const text = String(displayValue !== undefined && displayValue !== null && displayValue !== '' ? displayValue : value).trim();
  const ctx = String(context || '').trim();
  // Reject date-like text/ranges. These caused 2.420.260.000.000.700 style fake pack values.
  if (/\b(mon|tue|wed|thu|fri|sat|sun|senin|selasa|rabu|kamis|jumat|sabtu|minggu|jan|feb|mar|apr|may|mei|jun|jul|aug|agu|sep|oct|okt|nov|dec|des)\b/i.test(text)) return 0;
  if (/\d{1,2}\s*[-/]\s*\d{1,2}\s*[-/]\s*\d{2,4}/.test(text)) return 0;
  if (/\d{1,2}\s*[-/]\s*[A-Za-z]{3,}/.test(text)) return 0;
  if (/\b20\d{2}\b/.test(text) && !/rp|idr|porsi|pack|pax|qty|jumlah/i.test(text)) return 0;
  if (typeof value === 'number' && value > 30000) return 0; // likely Google Sheets date serial
  const n = kodAsNumber_(value);
  if (!isFinite(n) || n <= 0) return 0;
  if (n > (KOD_SAFE_LIMITS.maxProductionPackPerRow || 3000)) return 0;
  if (ctx && kodLooksLikeProductionNoise_(ctx)) return 0;
  return Math.round(n * 100) / 100;
}
function kodIsSafeProductionRow_(r, tomorrowKey) {
  if (!r || !r.menu) return false;
  if (kodLooksLikeProductionNoise_(r.menu)) return false;
  if (r.detail && r.detail.date && String(r.detail.date) !== String(tomorrowKey)) return false;
  const pack = kodSafeProductionPack_(r.packNum, r.detail && r.detail.qty, r.menu);
  return !!pack;
}
function kodLooksLikeProductionNoise_(s) {
  s = String(s || '').trim();
  if (!s) return true;
  if (kodLooksLikeNoise_(s)) return true;
  if (/\b(mon|tue|wed|thu|fri|sat|sun|senin|selasa|rabu|kamis|jumat|sabtu|minggu)\b/i.test(s)) return true;
  if (/\d{1,2}\s*[-/]\s*[A-Za-z]{3,}/.test(s)) return true;
  if (/format stiker|list kurir|apartemen|unit\s+\d+|double protein|tanpa karbo/i.test(s)) return true;
  return false;
}
function kodComputePackFromQtySize_(qty, sizeText) { const multiplier = kodSizeMultiplier_(sizeText); return qty * multiplier; }
function kodSizeMultiplier_(sizeText) { const s = String(sizeText || '').toUpperCase(); if (/LARGE|\bLG\b/.test(s)) return 8; if (/MEDIUM|\bMD\b/.test(s)) return 6; if (/SMALL|\bSM\b/.test(s)) return 4; if (/\bXS\b/.test(s)) return 2; if (/TEAM/.test(s)) return 1.5; if (/THINWALL/.test(s)) return 1; if (/NASI\s*KOTAK|BOX/.test(s)) return 1; return 1; }
