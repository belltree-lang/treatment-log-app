/**
 * Billing entry points for Apps Script deployment.
 *
 * These helper functions wire the Get/Logic/Output layers together so that
 * callers can preview JSON or generate patient invoice PDFs for a given billing month.
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('billing');
  template.baseUrl = ScriptApp.getService().getUrl() || '';
  return template
    .evaluate()
    .setTitle('請求処理アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Return the active spreadsheet (utility retained for compatibility).
 * @return {SpreadsheetApp.Spreadsheet}
 */
function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Resolve source data for billing generation (patients, visits, bank statuses).
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @return {Object} normalized source data including month metadata.
 */
function getBillingSource(billingMonth) {
  return getBillingSourceData(billingMonth);
}

const BILLING_CACHE_PREFIX = 'billing:prepared:';
const BILLING_CACHE_TTL_SECONDS = 21600; // 6 hours

function buildBillingCacheKey_(billingMonthKey) {
  const monthKey = String(billingMonthKey || '').trim();
  if (!monthKey) return '';
  return BILLING_CACHE_PREFIX + monthKey;
}

function getBillingCache_() {
  try {
    return CacheService.getScriptCache();
  } catch (err) {
    console.warn('[billing] CacheService unavailable', err);
    return null;
  }
}

function loadPreparedBilling_(billingMonthKey) {
  const key = buildBillingCacheKey_(billingMonthKey);
  if (!key) return null;
  const cache = getBillingCache_();
  if (!cache) return null;
  const cached = cache.get(key);
  if (!cached) return null;
  try {
    return JSON.parse(cached);
  } catch (err) {
    console.warn('[billing] Failed to parse prepared cache', err);
    return null;
  }
}

function savePreparedBilling_(payload) {
  const key = buildBillingCacheKey_(payload && payload.billingMonth);
  if (!key) return;
  const cache = getBillingCache_();
  if (!cache) return;
  try {
    cache.put(key, JSON.stringify(payload), BILLING_CACHE_TTL_SECONDS);
  } catch (err) {
    console.warn('[billing] Failed to cache prepared billing', err);
  }
}

function buildPreparedBillingPayload_(billingMonth) {
  const source = getBillingSourceData(billingMonth);
  const billingJson = generateBillingJsonFromSource(source);
  return {
    billingMonth: source.billingMonth,
    billingJson,
    preparedAt: new Date().toISOString(),
    patients: source.patients || source.patientMap || {},
    bankInfoByName: source.bankInfoByName || {},
    staffByPatient: source.staffByPatient || {},
    staffDirectory: source.staffDirectory || {},
    staffDisplayByPatient: source.staffDisplayByPatient || {}
  };
}

function coerceBillingJsonArray_(raw) {
  if (Array.isArray(raw)) return raw;
  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
      console.warn('[billing] Failed to parse billingJson string', err);
    }
  }
  return [];
}

function normalizePreparedBilling_(payload) {
  if (!payload) return null;
  return Object.assign({}, payload, { billingJson: coerceBillingJsonArray_(payload.billingJson) });
}

function toClientBillingPayload_(prepared) {
  const normalized = normalizePreparedBilling_(prepared);
  if (!normalized) return null;
  return {
    billingMonth: normalized.billingMonth || '',
    billingJson: normalized.billingJson,
    preparedAt: normalized.preparedAt || null,
    patients: normalized.patients || {},
    bankInfoByName: normalized.bankInfoByName || {},
    staffByPatient: normalized.staffByPatient || {},
    staffDirectory: normalized.staffDirectory || {},
    staffDisplayByPatient: normalized.staffDisplayByPatient || {}
  };
}

function prepareBillingData(billingMonth) {
  const prepared = buildPreparedBillingPayload_(billingMonth);
  savePreparedBilling_(prepared);
  return toClientBillingPayload_(prepared);
}

/**
 * Generate billing JSON without creating files (preview use case).
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @return {Object} billingMonth key and generated JSON array.
 */
function generateBillingJsonPreview(billingMonth) {
  const prepared = buildPreparedBillingPayload_(billingMonth);
  savePreparedBilling_(prepared);
  return prepared;
}

/**
 * Generate invoice PDFs and surface JSON for UI.
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @param {Object} [options] - Optional output overrides such as fileName or note.
 * @return {Object} billingMonth key, billingJson, and saved file metadata.
 */
function generateInvoices(billingMonth, options) {
  try {
    const prepared = buildPreparedBillingPayload_(billingMonth);
    savePreparedBilling_(prepared);
    return generatePreparedInvoices_(prepared, options);
  } catch (err) {
    const msg = err && err.message ? err.message : String(err);
    const stack = err && err.stack ? '\n' + err.stack : '';
    console.error('[generateInvoices] failed:', msg, stack);
    throw new Error('請求生成中にエラーが発生しました: ' + msg);
  }
}

function generatePreparedInvoices_(prepared, options) {
  const normalized = normalizePreparedBilling_(prepared);
  if (!normalized || !normalized.billingJson) {
    throw new Error('請求集計結果が見つかりません。先に集計を実行してください。');
  }
  const outputOptions = Object.assign({}, options, { billingMonth: normalized.billingMonth });
  const pdfs = generateInvoicePdfs(normalized.billingJson, outputOptions);
  const shouldExportBank = !outputOptions || outputOptions.skipBankExport !== true;
  const bankOutput = shouldExportBank ? exportBankTransferDataForPrepared_(normalized) : null;
  return {
    billingMonth: normalized.billingMonth,
    billingJson: normalized.billingJson,
    files: pdfs.files,
    bankOutput,
    preparedAt: normalized.preparedAt || null
  };
}

function generateInvoicesFromCache(billingMonth, options) {
  const month = normalizeBillingMonthInput(billingMonth);
  const prepared = normalizePreparedBilling_(loadPreparedBilling_(month.key));
  if (!prepared || !prepared.billingJson) {
    throw new Error('事前集計が見つかりません。先に「請求データを集計」ボタンを実行してください。');
  }
  return generatePreparedInvoices_(prepared, options);
}

function normalizeBillingEditBurden_(value) {
  if (value === null || value === undefined) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  if (raw === '自費') return '自費';
  const num = Number(raw);
  if (Number.isFinite(num) && num > 0) {
    if (num > 0 && num <= 1) return Math.round(num * 10);
    if (num < 10) return Math.round(num);
  }
  return null;
}

function normalizeBillingEditMedicalAssistance_(value) {
  if (value === undefined) return undefined;
  if (value === null) return 0;
  if (value === true) return 1;
  if (value === false) return 0;
  const num = Number(value);
  if (Number.isFinite(num)) return num ? 1 : 0;
  const text = String(value || '').trim().toLowerCase();
  if (!text) return 0;
  return ['1', 'true', 'yes', 'y', 'on', '有', 'あり', '〇', '○', '◯'].indexOf(text) >= 0 ? 1 : 0;
}

function normalizeBillingEdits_(maybeEdits) {
  if (!Array.isArray(maybeEdits)) return [];
  return maybeEdits.map(edit => {
    const pid = edit && edit.patientId ? String(edit.patientId).trim() : '';
    if (!pid) return null;
    const burden = normalizeBillingEditBurden_(edit.burdenRate);
    return {
      patientId: pid,
      insuranceType: edit.insuranceType != null ? String(edit.insuranceType).trim() : undefined,
      medicalAssistance: normalizeBillingEditMedicalAssistance_(edit.medicalAssistance),
      burdenRate: burden !== null ? burden : undefined,
      unitPrice: edit.unitPrice != null ? Number(edit.unitPrice) || 0 : undefined,
      carryOverAmount: edit.carryOverAmount != null ? Number(edit.carryOverAmount) || 0 : undefined,
      payerType: edit.payerType != null ? String(edit.payerType).trim() : undefined
    };
  }).filter(Boolean);
}

function applyBillingPatientEdits_(edits) {
  if (!Array.isArray(edits) || !edits.length) return { updated: 0 };
  const sheet = billingSs().getSheetByName(BILLING_PATIENT_SHEET_NAME);
  if (!sheet) {
    throw new Error('患者情報シートが見つかりません: ' + BILLING_PATIENT_SHEET_NAME);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { updated: 0 };
  const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const colPid = resolveBillingColumn_(headers, BILLING_LABELS.recNo, '患者ID', { required: true, fallbackIndex: BILLING_PATIENT_COLS_FIXED.recNo });
  const colInsurance = resolveBillingColumn_(headers, ['保険区分', '保険種別', '保険タイプ', '保険'], '保険区分', {});
  const colBurden = resolveBillingColumn_(headers, BILLING_LABELS.share, '負担割合', { fallbackIndex: BILLING_PATIENT_COLS_FIXED.share });
  const colUnitPrice = resolveBillingColumn_(headers, ['単価', '請求単価', '自費単価', '単価(自費)', '単価（自費）'], '単価', {});
  const colCarryOver = resolveBillingColumn_(headers, ['未入金', '未入金額', '未収金', '未収', '繰越', '繰越額', '繰り越し', '差引繰越', '前回未払', '前回未収', 'carryOverAmount'], '未入金額', {});
  const colPayer = resolveBillingColumn_(headers, ['保険者', '支払区分', '保険/自費', '保険区分種別'], '保険者', {});
  const colMedical = resolveBillingColumn_(headers, ['医療助成'], '医療助成', { fallbackLetter: 'AS' });

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const editMap = edits.reduce((map, edit) => {
    map[billingNormalizePatientId_(edit.patientId)] = edit;
    return map;
  }, {});
  const updates = [];

  values.forEach((row, idx) => {
    const pid = billingNormalizePatientId_(row[colPid - 1]);
    const edit = editMap[pid];
    if (!edit) return;
    const newRow = row.slice();
    if (colInsurance && edit.insuranceType !== undefined) newRow[colInsurance - 1] = edit.insuranceType;
    if (colMedical && edit.medicalAssistance !== undefined) newRow[colMedical - 1] = edit.medicalAssistance ? 1 : 0;
    if (colBurden && edit.burdenRate !== undefined) newRow[colBurden - 1] = edit.burdenRate;
    if (colUnitPrice && edit.unitPrice !== undefined) newRow[colUnitPrice - 1] = edit.unitPrice;
    if (colCarryOver && edit.carryOverAmount !== undefined) newRow[colCarryOver - 1] = edit.carryOverAmount;
    if (colPayer && edit.payerType !== undefined) newRow[colPayer - 1] = edit.payerType;
    updates.push({ rowNumber: idx + 2, values: newRow });
  });

  updates.forEach(update => {
    sheet.getRange(update.rowNumber, 1, 1, update.values.length).setValues([update.values]);
  });
  return { updated: updates.length };
}

function applyBillingEdits(billingMonth, options) {
  const opts = options || {};
  const edits = normalizeBillingEdits_(opts.edits);
  if (edits.length) {
    applyBillingPatientEdits_(edits);
  }
  const prepared = buildPreparedBillingPayload_(billingMonth);
  savePreparedBilling_(prepared);
  return prepared;
}

function applyBillingEditsAndGenerateInvoices(billingMonth, options) {
  const prepared = applyBillingEdits(billingMonth, options);
  return generatePreparedInvoices_(prepared, options || {});
}

function generateBankTransferData(billingMonth, options) {
  const prepared = buildPreparedBillingPayload_(billingMonth);
  savePreparedBilling_(prepared);
  return exportBankTransferDataForPrepared_(prepared, options || {});
}

function generateBankTransferDataFromCache(billingMonth, options) {
  const month = normalizeBillingMonthInput(billingMonth);
  const prepared = loadPreparedBilling_(month.key);
  if (!prepared || !prepared.billingJson) {
    throw new Error('事前集計が見つかりません。先に「請求データを集計」を実行してください。');
  }
  return exportBankTransferDataForPrepared_(prepared, options || {});
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
  return result;
}

function billingApplyPaymentResultsFromMenu() {
  const month = promptBillingMonthInput_();
  const result = applyBillingPaymentResultsEntry(month.key);
  SpreadsheetApp.getUi().alert('請求履歴を更新しました: ' + result.updated + ' 件');
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
  billingMenu.addItem('入金結果PDFの反映（アップロード→解析）', 'billingApplyPaymentResultPdfFromMenu');
  billingMenu.addItem('（管理者向け）履歴チェック', 'billingCheckHistoryFromMenu');
  billingMenu.addToUi();

  const attendanceMenu = ui.createMenu('勤怠管理');
  attendanceMenu.addItem('勤怠データを今すぐ同期', 'runVisitAttendanceSyncJobFromMenu');
  attendanceMenu.addItem('日次同期トリガーを確認', 'ensureVisitAttendanceSyncTriggerFromMenu');
  attendanceMenu.addToUi();
}
