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
  return { billingMonth: source.billingMonth, billingJson, bankJoinWarnings: summarizeBankJoinErrors_(billingJson) };
}

function summarizeBankJoinErrors_(billingJson) {
  const errors = (billingJson || []).filter(item => item && item.bankJoinError);
  return {
    count: errors.length,
    messages: errors.map(item => item.bankJoinMessage).filter(Boolean)
  };
}

function alertBankJoinWarnings_(billingJson) {
  const summary = summarizeBankJoinErrors_(billingJson);
  if (!summary.count) return summary;
  const ui = SpreadsheetApp.getUi();
  ui.alert([
    '銀行情報が紐づけられない患者が ' + summary.count + ' 名います。',
    '請求一覧の該当行を確認し、銀行シートを修正して再実行してください。'
  ].join('\n'));
  return summary;
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
  return { billingMonth: source.billingMonth, billingJson, excel: outputs.excel, csv: outputs.csv, history: outputs.history, bankJoinWarnings: summarizeBankJoinErrors_(billingJson) };
}

function generateBillingCsvOnlyEntry(billingMonth, options) {
  const source = getBillingSourceData(billingMonth);
  const billingJson = generateBillingJsonFromSource(source);
  const csv = createBillingCsvFile(billingJson, Object.assign({}, options, { billingMonth: source.billingMonth }));
  appendBillingHistoryRows(billingJson, { billingMonth: source.billingMonth, memo: options && options.note });
  return { billingMonth: source.billingMonth, billingJson, csv, bankJoinWarnings: summarizeBankJoinErrors_(billingJson) };
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

function extractDriveFileId_(input) {
  const raw = String(input || '').trim();
  if (!raw) return '';
  const match = raw.match(/[A-Za-z0-9_-]{25,}/);
  return match ? match[0] : '';
}

function promptBillingResultPdfFileId_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('入金結果PDFのURLまたはファイルIDを入力してください',
    '例: https://drive.google.com/file/d/XXXXX/view', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    throw new Error('入金結果PDF入力がキャンセルされました');
  }
  const fileId = extractDriveFileId_(response.getResponseText());
  if (!fileId) {
    throw new Error('Drive ファイルIDが特定できませんでした');
  }
  return fileId;
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
  alertBankJoinWarnings_(result.billingJson);
  return result;
}

function billingGenerateExcelFromMenu() {
  const month = promptBillingMonthInput_();
  const result = generateBillingOutputsEntry(month.key);
  SpreadsheetApp.getUi().alert('Excel/CSV を出力しました\nExcel: ' + result.excel.name + '\nCSV: ' + result.csv.name);
  alertBankJoinWarnings_(result.billingJson);
  return result;
}

function billingGenerateCsvFromMenu() {
  const month = promptBillingMonthInput_();
  const result = generateBillingCsvOnlyEntry(month.key);
  SpreadsheetApp.getUi().alert('CSV を出力しました: ' + result.csv.name);
  alertBankJoinWarnings_(result.billingJson);
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

function billingApplyPaymentResultPdfFromMenu() {
  const month = promptBillingMonthInput_();
  const fileId = promptBillingResultPdfFileId_();
  const result = applyBillingPaymentResultPdfEntry(month.key, fileId);
  SpreadsheetApp.getUi().alert([
    '入金結果PDFを解析しました。',
    '解析件数: ' + result.parsedCount,
    'マッチ件数: ' + result.matched,
    '履歴更新: ' + result.updated + ' 件'
  ].join('\n'));
  return result;
}

function summarizeVisitCounts_(counts) {
  const entries = Object.values(counts || {});
  const totalVisits = entries.reduce((sum, entry) => sum + (Number(entry && entry.visitCount) || 0), 0);
  return { patientCount: entries.length, totalVisits };
}

function generateVisitCountPreviewFromMenu() {
  const month = normalizeBillingMonthInput(new Date());
  const result = buildVisitCountMap_(month);
  const summary = summarizeVisitCounts_(result.counts);
  SpreadsheetApp.getUi().alert([
    '請求月: ' + result.billingMonth,
    '対象患者数: ' + summary.patientCount + ' 名',
    '合計施術回数: ' + summary.totalVisits + ' 回'
  ].join('\n'));
  return Object.assign({}, result, summary);
}

function billingGenerateJsonFromTreatmentLogs() {
  const month = normalizeBillingMonthInput(new Date());
  const result = generateBillingJsonPreview(month.key);
  SpreadsheetApp.getUi().alert('施術録を集計し、請求データを生成しました: ' + (result.billingJson ? result.billingJson.length : 0) + ' 件');
  alertBankJoinWarnings_(result.billingJson);
  return result;
}

function summarizeBillingHistory_(billingMonth) {
  const sheet = ss().getSheetByName('請求履歴');
  if (!sheet) {
    return { billingMonth, exists: false };
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { billingMonth, exists: true, total: 0, statuses: {}, paidTotal: 0, unpaidTotal: 0 };
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const rows = values.filter(row => row[0] === billingMonth);
  const statuses = {};
  let paidTotal = 0;
  let unpaidTotal = 0;
  rows.forEach(row => {
    const status = row[8] || '';
    statuses[status] = (statuses[status] || 0) + 1;
    paidTotal += Number(row[6]) || 0;
    unpaidTotal += Number(row[7]) || 0;
  });

  return { billingMonth, exists: true, total: rows.length, statuses, paidTotal, unpaidTotal };
}

function billingCheckHistoryFromMenu() {
  const month = promptBillingMonthInput_();
  const summary = summarizeBillingHistory_(month.key);
  const ui = SpreadsheetApp.getUi();
  if (!summary.exists) {
    ui.alert('請求履歴シートが存在しません。');
    return summary;
  }
  const statusTexts = Object.keys(summary.statuses || {}).filter(key => key)
    .map(key => key + ': ' + summary.statuses[key] + ' 件');
  const messageLines = [
    '請求月: ' + summary.billingMonth,
    '履歴件数: ' + summary.total,
    '入金合計: ' + summary.paidTotal.toLocaleString('ja-JP') + ' 円',
    '未入金合計: ' + summary.unpaidTotal.toLocaleString('ja-JP') + ' 円'
  ];
  if (statusTexts.length) {
    messageLines.push('ステータス内訳:');
    messageLines.push(statusTexts.join(', '));
  }
  ui.alert(messageLines.join('\n'));
  return summary;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const billingMenu = ui.createMenu('請求処理');
  billingMenu.addItem('請求データ生成（JSON）', 'billingGenerateJsonFromMenu');
  billingMenu.addItem('請求一覧Excel出力（テンプレ版）', 'billingGenerateExcelFromMenu');
  billingMenu.addItem('口座振替CSV出力', 'billingGenerateCsvFromMenu');
  billingMenu.addItem('入金結果PDFの反映（アップロード→解析）', 'billingApplyPaymentResultPdfFromMenu');
  billingMenu.addItem('合算請求書PDF出力（テンプレ版）', 'billingGenerateCombinedPdfsFromMenu');
  billingMenu.addItem('（管理者向け）履歴チェック', 'billingCheckHistoryFromMenu');
  billingMenu.addToUi();

  const attendanceMenu = ui.createMenu('勤怠管理');
  attendanceMenu.addItem('勤怠データを今すぐ同期', 'runVisitAttendanceSyncJobFromMenu');
  attendanceMenu.addItem('日次同期トリガーを確認', 'ensureVisitAttendanceSyncTriggerFromMenu');
  attendanceMenu.addToUi();
}
