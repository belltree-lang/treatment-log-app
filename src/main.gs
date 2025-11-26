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

function generateBillingCsvOnlyEntry(billingMonth, options) {
  const source = getBillingSourceData(billingMonth);
  const billingJson = generateBillingJsonFromSource(source);
  const csv = createBillingCsvFile(billingJson, Object.assign({}, options, { billingMonth: source.billingMonth }));
  appendBillingHistoryRows(billingJson, { billingMonth: source.billingMonth, memo: options && options.note });
  return { billingMonth: source.billingMonth, billingJson, csv };
}

function generateCombinedBillingPdfsEntry(billingMonth, options) {
  const source = getBillingSourceData(billingMonth);
  const billingJson = generateBillingJsonFromSource(source);
  const pdfs = generateCombinedBillingPdfs(billingJson, Object.assign({}, options, { billingMonth: source.billingMonth }));
  return { billingMonth: source.billingMonth, billingJson, pdfs };
}

function applyBillingPaymentResultsEntry(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const bankStatuses = getBillingPaymentResults(month);
  return applyPaymentResultsToHistory(month.key, bankStatuses);
}

function applyBillingPaymentResultPdfEntry(billingMonth, fileId, options) {
  const month = normalizeBillingMonthInput(billingMonth);
  const source = getBillingSourceData(month);
  const billingJson = generateBillingJsonFromSource(source);
  const file = DriveApp.getFileById(fileId);
  const pdfBlob = file.getBlob();
  const result = applyPaymentResultPdf(month.key, pdfBlob, billingJson);
  return Object.assign({}, result, {
    fileId,
    fileName: file.getName(),
    billingJson
  });
}

function promptBillingMonthInput_() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const defaultYm = Utilities.formatDate(today, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMM');
  const response = ui.prompt('請求月を入力してください (YYYYMM)', defaultYm, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    throw new Error('請求月入力がキャンセルされました');
  }
  return normalizeBillingMonthInput(response.getResponseText());
}

function billingGenerateJsonFromMenu() {
  const month = promptBillingMonthInput_();
  const result = generateBillingJsonPreview(month.key);
  SpreadsheetApp.getUi().alert('請求データを生成しました: ' + (result.billingJson ? result.billingJson.length : 0) + ' 件');
  return result;
}

function billingGenerateExcelFromMenu() {
  const month = promptBillingMonthInput_();
  const result = generateBillingOutputsEntry(month.key);
  SpreadsheetApp.getUi().alert('Excel/CSV を出力しました\nExcel: ' + result.excel.name + '\nCSV: ' + result.csv.name);
  return result;
}

function billingGenerateCsvFromMenu() {
  const month = promptBillingMonthInput_();
  const result = generateBillingCsvOnlyEntry(month.key);
  SpreadsheetApp.getUi().alert('CSV を出力しました: ' + result.csv.name);
  return result;
}

function billingApplyPaymentResultsFromMenu() {
  const month = promptBillingMonthInput_();
  const result = applyBillingPaymentResultsEntry(month.key);
  SpreadsheetApp.getUi().alert('請求履歴を更新しました: ' + result.updated + ' 件');
  return result;
}

function billingGenerateCombinedPdfsFromMenu() {
  const month = promptBillingMonthInput_();
  const result = generateCombinedBillingPdfsEntry(month.key);
  SpreadsheetApp.getUi().alert('合算請求書を ' + result.pdfs.count + ' 件作成しました');
  return result;
}
