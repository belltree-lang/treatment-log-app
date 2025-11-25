/**
 * Billing entry points for Apps Script deployment.
 *
 * These helper functions wire the Get/Logic/Output layers together so that
 * callers can preview JSON or generate Excel/CSV files for a given billing month.
 */

/**
 * Resolve source data for billing generation (patients, visits, bank statuses).
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @return {Object} normalized source data including month metadata.
 */
function getBillingSource(billingMonth) {
  return getBillingSourceData(billingMonth);
}

/**
 * Generate billing JSON without creating files (preview use case).
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @return {Object} billingMonth key and generated JSON array.
 */
function generateBillingJsonPreview(billingMonth) {
  const source = getBillingSourceData(billingMonth);
  const billingJson = generateBillingJsonFromSource(source);
  return { billingMonth: source.billingMonth, billingJson };
}

/**
 * Generate billing Excel/CSV outputs and record history.
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @param {Object} [options] - Optional overrides such as fileName or note.
 * @return {Object} billingMonth key, generated JSON, and output metadata.
 */
function generateBillingOutputsEntry(billingMonth, options) {
  const source = getBillingSourceData(billingMonth);
  const billingJson = generateBillingJsonFromSource(source);
  const outputOptions = Object.assign({}, options, { billingMonth: source.billingMonth });
  const outputs = generateBillingOutputs(billingJson, outputOptions);
  return { billingMonth: source.billingMonth, billingJson, excel: outputs.excel, csv: outputs.csv, history: outputs.history };
}
