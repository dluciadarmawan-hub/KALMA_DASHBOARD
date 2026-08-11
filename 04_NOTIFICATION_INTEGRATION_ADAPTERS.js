/**
 * Integration adapters to call AFTER successful save.
 * These functions do not save main transactions. They only notify/log after a save already succeeded.
 */
function notifyPurchasingApprovalAfterSave_(savedRecord) {
  return notifyOwnerIfNeeded_(Object.assign({ module: 'Purchasing' }, savedRecord || {}));
}

function notifyExpenseSubmissionAfterSave_(savedRecord) {
  return notifyOwnerIfNeeded_(Object.assign({ module: 'Purchasing', sourceSheet: 'EXPENSE_SUBMISSION' }, savedRecord || {}));
}

function notifyBcaChecklistAfterSave_(savedRecord) {
  return notifyOwnerIfNeeded_(Object.assign({ module: 'Purchasing', sourceSheet: 'BCA_PAYMENT_CHECKLIST' }, savedRecord || {}));
}

function notifyKasKecilSubmissionAfterSave_(savedRecord) {
  return notifyOwnerIfNeeded_(Object.assign({ module: 'Purchasing', sourceSheet: 'KAS_KECIL_SUBMISSION' }, savedRecord || {}));
}

function notifyPayrollApprovalAfterSave_(savedRecord) {
  return notifyOwnerIfNeeded_(Object.assign({ module: 'Payroll', sourceSheet: 'PAYROLL_APPROVAL' }, savedRecord || {}));
}

function notifyCashBankApprovalAfterSave_(savedRecord) {
  return notifyOwnerIfNeeded_(Object.assign({ module: 'Cash Bank' }, savedRecord || {}));
}

function notifyRevenueApprovalAfterSave_(savedRecord) {
  return notifyOwnerIfNeeded_(Object.assign({ module: 'Revenue' }, savedRecord || {}));
}

function notifyInvoiceApprovalAfterSave_(savedRecord) {
  return notifyOwnerIfNeeded_(Object.assign({ module: 'Invoice Engine' }, savedRecord || {}));
}

function notifyStockLogApprovalAfterSave_(savedRecord) {
  return notifyOwnerIfNeeded_(Object.assign({ module: 'Stock Log' }, savedRecord || {}));
}
