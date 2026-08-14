/**
 * KALMA_OWNER_DASHBOARD — standalone config.
 * READ ONLY. No sheet mutation. No posting. No sync.
 */
const KOD_BUILD = Object.freeze({
  name: 'OWNERDASH_20260814_V12',
  status: 'BUILD_CANDIDATE__V12_CASH_POSITION_VISIBILITY__READ_ONLY',
  revision: 'OWNERDASH_20260814_V12_CASH_POSITION_VISIBILITY_FIX',
  date: '2026-08-14',
  safety: 'READ_ONLY_OWNER_DASHBOARD__CENTRAL_AUTH_REQUIRED__NO_MASTER_OR_DOWNSTREAM_MUTATION'
});

const KOD_AUTH = Object.freeze({
  moduleCode: 'OWNER_DASHBOARD',
  authSource: 'CENTRAL_AUTH_LIBRARY_KalmaCore_V3',
  requiredLibraryCall: 'KalmaCore.requireStaffRole(pin, "OWNER_DASHBOARD", allowedRoles)',
  sessionCachePrefix: 'KOD_CENTRAL_SESSION_',
  sessionTtlSeconds: 43200,
  clientRoleTrustPolicy: 'NEVER_TRUST_LOCALSTORAGE_OR_CLIENT_ROLE_FOR_PERMISSION',
  legacyRoleAliases: Object.freeze({
    OPS: 'OPS_ADMIN'
  }),
  roleSets: Object.freeze({
    login: Object.freeze(['OWNER', 'OPS_ADMIN', 'STAFF']),
    dashboardRead: Object.freeze(['OWNER', 'OPS_ADMIN', 'STAFF']),
    staffLevel: Object.freeze(['STAFF', 'OPS_ADMIN']),
    ownerOnly: Object.freeze(['OWNER'])
  })
});


const KOD_SOURCE_IDS = Object.freeze({
  MASTER_SCHEDULE: '1sEXhuvtcCEtXEcLxpKyaXjVVN0cXxRtSnh0VNwnG_l4',
  PURCHASING: '1ho1LX4UbD6HHxUkK_qi_vmbcx2rMe-rmjnrugiYJrpE',
  CASH_BANK: '1rd4-7K9VKjyvxalmMsibNzm7aq76xRqMxSaNeu1HQO0',
  REVENUE: '1pLTf4pZFHcu26UBmkbH3hLJFNQpO4eH-Y8wO37XveCk',
  STOCK_LOG: '1A9wL-FWP7t9OvlhzVgbgKmSwo5b0Uh5pfxOup4WPaok',
  INVOICE_CUSTOMER: '1kQAfXGZUdq2DIc7xIsnDge3KS1wNlZQCVMgSpyKvvlI',
  BILLING_RESERVE: '1OdoMoc-i5tzDHOjMFw94sB_8-UdrLP96ctuH-TyKTZg',
  APPS_SCRIPT_PROJECTS: '1ORz448m4lEcHtjKpDtD6ezJusg8OOlkPs5tPIpv_KnA',
  ORDER_BOARD: '1YL8kxInf0GMPUN2E57rksgWDb8C7Og7P4f2yGcVFxcM',
  RECEIVING: '1RhKaW-72F4CSLnYcqTiAYcHj6U9R7CrrDznqD0OmABY'
});

const KOD_ROUTE_LINKS = Object.freeze({
  MASTER_SCHEDULE_SHEET: 'https://docs.google.com/spreadsheets/d/' + KOD_SOURCE_IDS.MASTER_SCHEDULE,
  PURCHASING_SHEET: 'https://docs.google.com/spreadsheets/d/' + KOD_SOURCE_IDS.PURCHASING,
  CASH_BANK_SHEET: 'https://docs.google.com/spreadsheets/d/' + KOD_SOURCE_IDS.CASH_BANK,
  REVENUE_SHEET: 'https://docs.google.com/spreadsheets/d/' + KOD_SOURCE_IDS.REVENUE,
  STOCK_LOG_SHEET: 'https://docs.google.com/spreadsheets/d/' + KOD_SOURCE_IDS.STOCK_LOG,
  INVOICE_CUSTOMER_SHEET: 'https://docs.google.com/spreadsheets/d/' + KOD_SOURCE_IDS.INVOICE_CUSTOMER,
  BILLING_RESERVE_SHEET: 'https://docs.google.com/spreadsheets/d/' + KOD_SOURCE_IDS.BILLING_RESERVE,
  APPS_SCRIPT_PROJECTS_SHEET: 'https://docs.google.com/spreadsheets/d/' + KOD_SOURCE_IDS.APPS_SCRIPT_PROJECTS,
  ORDER_BOARD_SHEET: 'https://docs.google.com/spreadsheets/d/' + KOD_SOURCE_IDS.ORDER_BOARD,
  RECEIVING_SHEET: 'https://docs.google.com/spreadsheets/d/' + KOD_SOURCE_IDS.RECEIVING,

  MASTER_SCHEDULE_WEBAPP: 'https://script.google.com/macros/s/AKfycbwwGH7OZ3aF3pG8-ir45gCVzDbEXuIlGVaYxIe4m_NJrBZ0k_ewSl4IufNVai0LSE9Yhw/exec',
  PURCHASING_WEBAPP: 'https://script.google.com/macros/s/AKfycbw0GMKgn93HjoKZw-WXlaouYfzqGMHqLZjgBHj8hgUc2WGjaXZsM7AH5Tb0wffT5235zw/exec',
  CASH_BANK_WEBAPP: 'https://script.google.com/macros/s/AKfycbw-nh1CdUv4uPF1P5OnYvl04Ngf9-9K38_9QzaXzryZNWe8g3ER49Sy2tsmWH3U4BzA/exec',
  REVENUE_WEBAPP: 'https://script.google.com/macros/s/AKfycbwz5PW6S6BryuBIWVsyfxvsTwxa_TQyXGnFvLmGD9b4VY05aC_ejOBwetuK78WzGaDV/exec',
  STOCK_LOG_WEBAPP: 'https://script.google.com/macros/s/AKfycbxpHl98yao_9tR0g8sr4z7zlzqvlrXqd9yKxIDHOT9te4xjK_sFllZ4DvjxZ4uOP8U/exec',
  PAYROLL_WEBAPP: 'https://script.google.com/macros/s/AKfycbw24qeQz94Z7erH6KJVY5cODXZ8UE3vkOyYFx0ODvSbvMCYV1KrwblwmyJYeFolFSKw/exec',
  INVOICE_WEBAPP: 'https://script.google.com/macros/s/AKfycbz8CRnMHKJn4qzXg8Vta65XAPTSgWTlKx9FzfPHj3jeD87SmIhJdgfCgL9c4SGun1VYVg/exec',
  RECEIVING_WEBAPP: 'https://script.google.com/macros/s/AKfycbx0JP6wbBcQD0fu6oj-AGj3XzxFNehGeekit-d92fWWCSnCqrOJHekANrc9_o0fIpym/exec',
  ORDER_BOARD_WEBAPP: 'https://script.google.com/macros/s/AKfycbw19TK5zDhcFFvSq6TK2KtQQk3-sjYk-yfr4Bw8Y47JVcBPBollOjAuRXS0U-SoXAfq/exec',
  OWNER_DASHBOARD_WEBAPP: 'https://script.google.com/macros/s/AKfycbwDNGuXANoArHEPZYIGecYVyR-s8Y-r0rsgE2M_hTvks7y0QfmUz9VXag0uI2Jv6EfH/exec'
});

const KOD_SAFE_LIMITS = Object.freeze({
  maxRowsPerSheet: 900,
  maxColsPerSheet: 40,
  statusScanMaxRows: 1400,
  maxDashboardRowsPerPanel: 8,
  cashLedgerMaxRows: 5000,
  cashboxMovementMaxRows: 3000,
  maxProductionPackPerRow: 3000,
  maxProductionTotalPack: 15000
});



const KOD_MASTER_APPROVAL_R1 = Object.freeze({
  sourceSystem: 'ORDER_BOARD',
  sourceBuild: 'ORDERBOARD_20260627_V25',
  sourceSha256: '1e918f251fcbaadfc1f2c2ed4e7bf35109d0c997194b8a008812e2e975810cc0',
  queueSheet: 'INBOUND_MASTER_APPROVAL_QUEUE',
  dashboardBuild: 'OWNERDASH_20260814_V12',
  allowedReadOnlyColumns: [
    'Candidate_ID', 'Candidate_Type', 'Candidate_Status', 'Raw_Typed_Name', 'Normalized_Name', 'Source_Module',
    'First_Order_ID', 'First_Order_Line_ID', 'Source_Order_IDs', 'Source_Order_Line_IDs',
    'Vendor_Candidate_ID', 'Item_Candidate_ID', 'Vendor_Master_ID', 'Vendor_Resolved_Name',
    'Master_Item_ID', 'Item_Resolved_Name', 'Item_Resolved_Unit',
    'Source_Order_ID', 'Source_Line_ID', 'Status', 'Created_At', 'Created_By',
    'Resolved_At', 'Resolved_By', 'Suggested_Action'
  ],
  disabledVendorActions: [
    'APPROVE_AS_NEW_VENDOR', 'MAP_TO_EXISTING_VENDOR', 'REJECT_INVALID', 'MERGE_DUPLICATE'
  ],
  disabledItemActions: [
    'APPROVE_AS_NEW_ITEM', 'MAP_TO_EXISTING_ITEM', 'REJECT_INVALID', 'MERGE_DUPLICATE'
  ],
  closedStatusRegex: '^(APPROVED|RESOLVED|REJECTED|MERGED|MAPPED|CLOSED|VOIDED|DONE|COMPLETED)$'
});

const KOD_FINAL_STATUSES = Object.freeze([
  'PAID', 'PAID_FULL', 'FULLY_PAID', 'CLOSED', 'CANCELLED', 'VOIDED', 'REJECTED_FINAL',
  'SYNCED CLEAN', 'SYNCED_CLEAN', 'FINAL_CLOSED', 'DONE', 'COMPLETED'
]);
