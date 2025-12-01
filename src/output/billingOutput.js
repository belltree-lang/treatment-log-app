/***** Output layer: billing invoice PDF generation *****/

const INVOICE_PARENT_FOLDER_ID = '1EG-GB3PbaUr9C1LJWlaf_idqoYF-19Ux';
const INVOICE_FILE_PREFIX = '請求書';
const TRANSPORT_PRICE = (typeof BILLING_TRANSPORT_UNIT_PRICE !== 'undefined')
  ? BILLING_TRANSPORT_UNIT_PRICE
  : 33;
const INVOICE_TREATMENT_UNIT_PRICE_BY_BURDEN = { 1: 417, 2: 834, 3: 1251 };

function roundToNearestTen_(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.round(num / 10) * 10;
}

function escapeHtml_(value) {
  return String(value || '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[c]);
}

const normalizeInvoiceBurdenRateInt_ = typeof normalizeBurdenRateInt_ === 'function'
  ? normalizeBurdenRateInt_
  : function fallbackNormalizeInvoiceBurdenRateInt_(burdenRate) {
    if (burdenRate == null || burdenRate === '') return 0;
    const num = Number(burdenRate);
    if (Number.isFinite(num)) {
      if (num > 0 && num < 1) return Math.round(num * 10);
      if (num >= 1 && num < 10) return Math.round(num);
      if (num >= 10 && num <= 100) return Math.round(num / 10);
    }
    const normalized = String(burdenRate).normalize('NFKC').replace(/\s+/g, '').replace('％', '%');
    const withoutUnits = normalized.replace(/割|分/g, '').replace('%', '');
    const parsed = Number(withoutUnits);
    if (!Number.isFinite(parsed)) return 0;
    if (normalized.indexOf('%') >= 0) return Math.round(parsed / 10);
    if (parsed > 0 && parsed < 10) return Math.round(parsed);
    if (parsed >= 10 && parsed <= 100) return Math.round(parsed / 10);
    return 0;
  };

function convertSpreadsheetToExcelBlob_(file, exportName) {
  if (!file || typeof file.getMimeType !== 'function') {
    throw new Error('スプレッドシート以外のファイルをExcelに変換することはできません');
  }

  const mimeType = file.getMimeType();
  const isSpreadsheet = mimeType === MimeType.GOOGLE_SHEETS;
  const isExcel = mimeType === MimeType.MICROSOFT_EXCEL;
  if (!isSpreadsheet && !isExcel) {
    throw new Error('スプレッドシート以外のファイルをExcelに変換することはできません');
  }

  const blob = file.getBlob();
  const name = (exportName && String(exportName).trim()) || 'export';
  const excelBlob = isSpreadsheet && typeof blob.getAs === 'function'
    ? blob.getAs(MimeType.MICROSOFT_EXCEL)
    : blob;
  return excelBlob.setName(name + '.xlsx');
}

function normalizeBillingAmount_(item) {
  if (!item) return 0;

  if (item.grandTotal != null && item.grandTotal !== '') {
    return normalizeInvoiceMoney_(item.grandTotal);
  }

  const carryOverTotal = normalizeInvoiceMoney_(item.carryOverAmount)
    + normalizeInvoiceMoney_(item.carryOverFromHistory);

  if (item.total != null && item.total !== '') {
    return normalizeInvoiceMoney_(item.total) + carryOverTotal;
  }

  const billingAmount = normalizeInvoiceMoney_(item.billingAmount);
  const treatmentAmount = normalizeInvoiceMoney_(item.treatmentAmount);
  const transportAmount = normalizeInvoiceMoney_(item.transportAmount);

  if (item.billingAmount != null && item.billingAmount !== '') {
    return billingAmount + transportAmount + carryOverTotal;
  }

  if (item.treatmentAmount != null || item.transportAmount != null || carryOverTotal) {
    return treatmentAmount + transportAmount + carryOverTotal;
  }

  return 0;
}

function normalizeBillingNameKey_(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}

function normalizeInvoiceMoney_(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }
  const text = String(value || '')
    .normalize('NFKC')
    .replace(/[，,]/g, '')
    .trim();
  if (!text) return 0;
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
}

function normalizeInvoiceVisitCount_(value) {
  const source = value && value.visitCount != null ? value.visitCount : value;
  if (typeof source === 'number') {
    return Number.isFinite(source) && source > 0 ? source : 0;
  }
  const normalized = String(source || '')
    .normalize('NFKC')
    .replace(/[，,]/g, '')
    .trim();
  const num = Number(normalized);
  return Number.isFinite(num) && num > 0 ? num : 0;
}

function normalizeInvoiceMedicalAssistanceFlag_(value) {
  if (value === true) return true;
  if (value === false) return false;
  const num = Number(value);
  if (Number.isFinite(num)) return !!num;
  const text = String(value || '').trim().toLowerCase();
  if (!text) return false;
  return ['1', 'true', 'yes', 'y', 'on', '有', 'あり', '〇', '○', '◯'].indexOf(text) >= 0;
}

function normalizeBillingCarryOver_(item) {
  if (!item) return 0;
  const directCarryOver = (item.carryOverAmount != null && item.carryOverAmount !== '')
    ? normalizeInvoiceMoney_(item.carryOverAmount)
    : (item.raw && item.raw.carryOverAmount != null)
      ? normalizeInvoiceMoney_(item.raw.carryOverAmount)
      : 0;
  const historyCarryOver = normalizeInvoiceMoney_(item.carryOverFromHistory);
  return directCarryOver + historyCarryOver;
}

function formatBillingCurrency_(value) {
  const num = Number(value);
  if (!isFinite(num)) return '0';
  return Math.round(num).toLocaleString('ja-JP');
}

function normalizeBillingMonthLabel_(billingMonth) {
  const digits = (billingMonth ? String(billingMonth) : '').replace(/\D/g, '');
  if (digits.length >= 6) {
    const year = digits.slice(0, 4);
    const month = digits.slice(4, 6).padStart(2, '0');
    return year + '年' + month + '月';
  }
  return billingMonth || '';
}

function formatBillingMonthForFile_(billingMonth) {
  const digits = (billingMonth ? String(billingMonth) : '').replace(/\D/g, '');
  if (digits.length >= 6) {
    const year = digits.slice(0, 4);
    const month = digits.slice(4, 6).padStart(2, '0');
    return year + '-' + month;
  }
  return billingMonth || '';
}

function formatBillingMonthCompact_(billingMonth) {
  const digits = (billingMonth ? String(billingMonth) : '').replace(/\D/g, '');
  if (digits.length >= 6) {
    return digits.slice(0, 6);
  }
  return '';
}

function formatInvoiceDateLabel_() {
  try {
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    return Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  } catch (e) {
    return '';
  }
}

function sanitizeFileName_(text) {
  const raw = String(text || '').trim();
  return raw ? raw.replace(/[\\/\r\n]/g, '_') : '請求書';
}

function formatInvoiceFileName_(item) {
  const baseName = sanitizeFileName_(item && (item.nameKanji || item.patientId || INVOICE_FILE_PREFIX));
  const dateLabel = formatInvoiceDateLabel_();
  return baseName + '_' + (dateLabel || 'YYYYMMDD') + '_請求書.pdf';
}

function buildInvoiceTemplateData_(item) {
  const billingMonth = item && item.billingMonth;
  const monthLabel = normalizeBillingMonthLabel_(billingMonth);
  const breakdown = calculateInvoiceChargeBreakdown_(Object.assign({}, item, { billingMonth }));
  const visits = breakdown.visits || 0;
  const unitPrice = breakdown.treatmentUnitPrice || 0;
  const rows = [
    { label: '前月繰越', detail: '', amount: normalizeBillingCarryOver_(item) },
    { label: '施術料', detail: formatBillingCurrency_(unitPrice) + '円 × ' + visits + '回', amount: breakdown.treatmentAmount },
    { label: '交通費', detail: formatBillingCurrency_(TRANSPORT_PRICE) + '円 × ' + visits + '回', amount: breakdown.transportAmount }
  ];

  return Object.assign({}, item, {
    monthLabel,
    rows,
    grandTotal: breakdown.grandTotal
  });
}

function createInvoicePdfBlob_(item) {
  const template = HtmlService.createTemplateFromFile('invoice_template');
  template.data = buildInvoiceTemplateData_(item || {});
  const html = template.evaluate().setWidth(1240).setHeight(1754);
  const fileName = formatInvoiceFileName_(item);
  return html.getBlob().getAs(MimeType.PDF).setName(fileName);
}

function ensureInvoiceRootFolder_() {
  if (!INVOICE_PARENT_FOLDER_ID) {
    throw new Error('請求書の保存先フォルダIDが設定されていません');
  }
  return DriveApp.getFolderById(INVOICE_PARENT_FOLDER_ID);
}

function ensureSubFolder_(parentFolder, name) {
  const folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(name);
}

function formatResponsibleFolderName_(billingMonth, responsibleName) {
  const ym = formatBillingMonthCompact_(billingMonth);
  const ymLabel = ym || '請求月未設定';
  const safeName = sanitizeFileName_(responsibleName || '担当者未設定');
  return ymLabel + '請求書_' + safeName;
}

function calculateInvoiceChargeBreakdown_(params) {
  const visits = normalizeInvoiceVisitCount_(params && params.visitCount);
  const insuranceType = params && params.insuranceType ? String(params.insuranceType).trim() : '';
  const burdenRateInt = normalizeInvoiceBurdenRateInt_(params && params.burdenRate);
  const carryOverAmount = normalizeBillingCarryOver_(params);
  const manualUnitPrice = normalizeInvoiceMoney_(params && params.unitPrice);
  const hasManualUnitPrice = Number.isFinite(manualUnitPrice) && manualUnitPrice !== 0;
  const isMedicalAssistance = normalizeInvoiceMedicalAssistanceFlag_(params && params.medicalAssistance);
  const isMassage = insuranceType === 'マッサージ';
  const isSelfPaid = insuranceType === '自費';
  const shouldZero = (insuranceType === '生保' || isMedicalAssistance) && !hasManualUnitPrice;
  const isZeroChargeInsurance = shouldZero || (insuranceType === '自費' && !hasManualUnitPrice);

  const treatmentUnitPrice = (function resolveTreatmentUnitPrice() {
    if (shouldZero) return 0;
    if (isMassage) return 0;
    if (hasManualUnitPrice) return manualUnitPrice;
    if (insuranceType === '自費') return 0;
    return INVOICE_TREATMENT_UNIT_PRICE_BY_BURDEN[burdenRateInt] || 0;
  })();
  const rawTreatmentAmount = visits > 0 ? treatmentUnitPrice * visits : 0;
  const treatmentAmount = (isZeroChargeInsurance || isMassage)
    ? rawTreatmentAmount
    : (isSelfPaid ? rawTreatmentAmount : roundToNearestTen_(rawTreatmentAmount));
  const transportAmount = visits > 0 && !isMassage && !isZeroChargeInsurance
    ? TRANSPORT_PRICE * visits
    : 0;
  const grandTotal = carryOverAmount + treatmentAmount + transportAmount;

  return { treatmentUnitPrice, treatmentAmount, transportAmount, grandTotal, visits };
}

  function buildBillingInvoiceHtml_(item, billingMonth) {
    const targetMonth = billingMonth || (item && item.billingMonth) || '';
    const breakdown = calculateInvoiceChargeBreakdown_(Object.assign({}, item, { billingMonth: targetMonth }));
    const monthLabel = normalizeBillingMonthLabel_(targetMonth);
    const visits = breakdown.visits || 0;
  const treatmentUnitPrice = breakdown.treatmentUnitPrice || 0;
  const transportUnitPrice = TRANSPORT_PRICE;
  const carryOverAmount = normalizeBillingCarryOver_(item);
  const totalLabel = formatBillingCurrency_(breakdown.grandTotal) + '円';

    const name = escapeHtml_((item && item.nameKanji) || '');

    return [
      '<div class="billing-invoice">',
      '<h1>べるつりー訪問鍼灸マッサージ</h1>',
      `<h2>${escapeHtml_(monthLabel)} ご請求書</h2>`,
      name ? `<p class="patient-name">${name} 様</p>` : '',
      '<div class="charge-breakdown">',
      `<p>前月繰越: ${formatBillingCurrency_(carryOverAmount)}円</p>`,
    `<p>施術料（${formatBillingCurrency_(treatmentUnitPrice)}円 × ${visits}回）: ${formatBillingCurrency_(breakdown.treatmentAmount)}円</p>`,
    `<p>交通費（${formatBillingCurrency_(transportUnitPrice)}円 × ${visits}回）: ${formatBillingCurrency_(breakdown.transportAmount)}円</p>`,
    `<p class="grand-total">合計: ${totalLabel}</p>`,
    '</div>',
    '</div>'
  ].filter(Boolean).join('');
}

function ensureInvoiceFolderForResponsible_(item) {
  const root = ensureInvoiceRootFolder_();
  const folderName = formatResponsibleFolderName_(item && item.billingMonth, item && item.responsibleName);
  return ensureSubFolder_(root, folderName);
}

function removeExistingInvoiceFiles_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    const file = files.next();
    try {
      file.setTrashed(true);
    } catch (e) {
      // ignore cleanup errors
    }
  }
}

function saveInvoicePdf(item, pdfBlob) {
  const folder = ensureInvoiceFolderForResponsible_(item);
  const fileName = pdfBlob.getName();
  removeExistingInvoiceFiles_(folder, fileName);
  const file = folder.createFile(pdfBlob);
  return { fileId: file.getId(), url: file.getUrl(), name: file.getName() };
}

function generateInvoicePdf(item) {
  const blob = createInvoicePdfBlob_(item);
  return saveInvoicePdf(item, blob);
}

function generateInvoicePdfs(billingJson, options) {
  const billingMonth = (options && options.billingMonth) || (Array.isArray(billingJson) && billingJson.length && billingJson[0].billingMonth) || '';
  const files = (billingJson || []).map(item => {
    const meta = generateInvoicePdf(item);
    return Object.assign({}, meta, { patientId: item && item.patientId, nameKanji: item && item.nameKanji });
  });
  return { billingMonth, files };
}

/***** Bank transfer export helpers *****/

const BANK_TRANSFER_HEADERS = ['請求月', '番号', '氏名（漢字）', '銀行コード', '支店コード', '規定コード', '口座番号', '氏名（カナ）', '新規フラグ'];

function buildBankTransferRowsForBilling_(billingJson, bankInfoByName, patientMap, billingMonth) {
  const rows = [];
  let skipped = 0;
  const billingMonthKey = billingMonth || (billingJson && billingJson.length ? billingJson[0].billingMonth : '');

  (billingJson || []).forEach(item => {
    const pid = item && item.patientId ? String(item.patientId).trim() : '';
    if (!pid) return;
    const patient = patientMap && patientMap[pid] ? patientMap[pid] : {};
    const nameKanji = item && item.nameKanji ? String(item.nameKanji).trim() : '';
    const nameKey = normalizeBillingNameKey_(nameKanji);
    const bankLookup = bankInfoByName && nameKey ? bankInfoByName[nameKey] : null;

    const bankCode = (bankLookup && bankLookup.bankCode) || patient.bankCode || '';
    const branchCode = (bankLookup && bankLookup.branchCode) || patient.branchCode || '';
    const regulationCode = (bankLookup && bankLookup.regulationCode) || patient.regulationCode || 1;
    const accountNumber = (bankLookup && bankLookup.accountNumber) || patient.accountNumber || '';
    const nameKana = (bankLookup && bankLookup.nameKana) || patient.nameKana || (item && item.nameKana) || '';
    const isNew = normalizeZeroOneFlag_((bankLookup && bankLookup.isNew) != null ? bankLookup.isNew : patient.isNew);

    if (!bankCode || !branchCode || !accountNumber) {
      skipped += 1;
      return;
    }

    rows.push({
      billingMonth: billingMonthKey,
      patientId: pid,
      nameKanji,
      bankCode,
      branchCode,
      regulationCode,
      accountNumber,
      nameKana,
      isNew
    });
  });

  return { billingMonth: billingMonthKey, rows, skipped };
}

function ensureBankTransferSheet_() {
  const workbook = billingSs();
  let sheet = workbook.getSheetByName(BILLING_BANK_SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(BILLING_BANK_SHEET_NAME);
    sheet.getRange(1, 1, 1, BANK_TRANSFER_HEADERS.length).setValues([BANK_TRANSFER_HEADERS]);
    return { sheet, headers: BANK_TRANSFER_HEADERS.slice() };
  }

  const lastCol = Math.max(sheet.getLastColumn(), BANK_TRANSFER_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  return { sheet, headers };
}

function resolveBankTransferColumns_(sheet, headers) {
  const workingHeaders = headers ? headers.slice() : [];
  const resolved = {};

  function ensureColumn(label, candidates) {
    const idx = resolveBillingColumn_(workingHeaders, candidates, label, {});
    if (idx) {
      resolved[label] = idx;
      return idx;
    }
    const newIndex = workingHeaders.length + 1;
    sheet.getRange(1, newIndex).setValue(label);
    workingHeaders.push(label);
    resolved[label] = newIndex;
    return newIndex;
  }

  ensureColumn('請求月', ['請求月', 'billingMonth', '請求年月']);
  ensureColumn('番号', BILLING_LABELS.recNo.concat(['番号', '患者番号', '患者ID']));
  ensureColumn('氏名（漢字）', BILLING_LABELS.name.concat(['氏名', '氏名（漢字）']));
  ensureColumn('銀行コード', ['銀行コード', '銀行CD', '銀行番号', 'bankCode']);
  ensureColumn('支店コード', ['支店コード', '支店番号', '支店CD', 'branchCode']);
  ensureColumn('規定コード', ['規定コード', '規定', '規定CD', '規定コード(1固定)', '1固定']);
  ensureColumn('口座番号', ['口座番号', '口座No', '口座NO', 'accountNumber', '口座']);
  ensureColumn('氏名（カナ）', BILLING_LABELS.furigana.concat(['氏名（カナ）']));
  ensureColumn('新規フラグ', ['新規', '新患', 'isNew', '新規フラグ', '新規区分']);

  return { columns: resolved, headers: workingHeaders };
}

function exportBankTransferRows_(billingMonth, rowObjects) {
  const ensured = ensureBankTransferSheet_();
  const sheet = ensured.sheet;
  const { columns, headers } = resolveBankTransferColumns_(sheet, ensured.headers);
  const colCount = Math.max(sheet.getLastColumn(), headers.length, Math.max.apply(null, Object.values(columns)));

  const lastRow = sheet.getLastRow();
  const existingValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, colCount).getValues() : [];
  const filtered = columns['請求月']
    ? existingValues.filter(row => String(row[columns['請求月'] - 1] || '').trim() !== String(billingMonth))
    : existingValues;

  const mapped = (rowObjects || []).map(obj => {
    const row = new Array(colCount).fill('');
    row[columns['請求月'] - 1] = billingMonth;
    row[columns['番号'] - 1] = obj.patientId || '';
    row[columns['氏名（漢字）'] - 1] = obj.nameKanji || '';
    row[columns['銀行コード'] - 1] = obj.bankCode || '';
    row[columns['支店コード'] - 1] = obj.branchCode || '';
    row[columns['規定コード'] - 1] = obj.regulationCode || '';
    row[columns['口座番号'] - 1] = obj.accountNumber || '';
    row[columns['氏名（カナ）'] - 1] = obj.nameKana || '';
    row[columns['新規フラグ'] - 1] = obj.isNew || '';
    return row;
  });

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, colCount).clearContent();
  }
  const newRows = filtered.concat(mapped);
  if (newRows.length) {
    sheet.getRange(2, 1, newRows.length, colCount).setValues(newRows);
  }

  return { billingMonth, inserted: mapped.length };
}

function exportBankTransferDataForPrepared_(prepared) {
  if (!prepared || !prepared.billingJson) {
    throw new Error('請求データが見つかりません。先に集計を実行してください。');
  }
  let bankInfoByName = prepared.bankInfoByName || {};
  let patientMap = prepared.patients || prepared.patientMap || {};
  if (!Object.keys(bankInfoByName).length || !Object.keys(patientMap).length) {
    const source = getBillingSourceData(prepared.billingMonth);
    bankInfoByName = source.bankInfoByName || bankInfoByName;
    patientMap = source.patients || source.patientMap || patientMap;
    if (!prepared.billingJson || !prepared.billingJson.length) {
      prepared.billingJson = generateBillingJsonFromSource(source);
    }
  }
  const buildResult = buildBankTransferRowsForBilling_(prepared.billingJson, bankInfoByName, patientMap, prepared.billingMonth);
  const outputResult = exportBankTransferRows_(buildResult.billingMonth, buildResult.rows);
  return Object.assign({}, buildResult, outputResult);
}

/***** Billing history and payment result utilities (retained for compatibility) *****/

function ensureBillingHistorySheet_() {
  const SHEET_NAME = '請求履歴';
  const workbook = ss();
  let sheet = workbook.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 11).setValues([[
      'billingMonth',
      'patientId',
      'nameKanji',
      'billingAmount',
      'carryOverAmount',
      'grandTotal',
      'paidAmount',
      'unpaidAmount',
      'bankStatus',
      'updatedAt',
      'memo'
    ]]);
  }
  return sheet;
}

function appendBillingHistoryRows(billingJson, options) {
  const opts = options || {};
  const sheet = ensureBillingHistorySheet_();
  const billingMonth = opts.billingMonth || (Array.isArray(billingJson) && billingJson.length && billingJson[0].billingMonth) || '';
  const memo = opts.memo || '';
  const rows = (billingJson || []).map(item => {
    const billingAmount = item && item.treatmentAmount != null ? item.treatmentAmount : 0;
    const carryOver = item && item.carryOverAmount != null ? item.carryOverAmount : 0;
    const paid = item && item.paidAmount != null ? item.paidAmount : 0;
    const grandTotal = normalizeBillingAmount_(item);
    const unpaid = grandTotal - paid;
    return [
      billingMonth,
      item && item.patientId ? item.patientId : '',
      item && item.nameKanji ? item.nameKanji : '',
      billingAmount,
      carryOver,
      grandTotal,
      paid,
      unpaid,
      item && item.bankStatus ? item.bankStatus : '',
      new Date(),
      memo
    ];
  });
  if (rows.length) {
    sheet.insertRows(2, rows.length);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  return { billingMonth, inserted: rows.length };
}

function applyPaymentResultsToHistory(billingMonth, bankStatuses) {
  const sheet = ensureBillingHistorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { billingMonth, updated: 0 };
  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const updates = [];
  const statusMap = bankStatuses || {};
  data.forEach((row, idx) => {
    if (row[0] !== billingMonth) return;
    const pid = row[1];
    const statusEntry = statusMap[pid];
    if (!statusEntry) return;
    const newRow = row.slice();
    let changed = false;

    if (statusEntry.bankStatus) {
      newRow[8] = statusEntry.bankStatus;
      changed = true;
    }
    if (statusEntry.paidAmount != null) {
      const paid = Number(statusEntry.paidAmount) || 0;
      newRow[6] = paid;
      const grandTotal = Number(newRow[5]) || 0;
      const unpaid = statusEntry.unpaidAmount != null ? statusEntry.unpaidAmount : grandTotal - paid;
      newRow[7] = unpaid;
      changed = true;
    }

    if (!changed) return;
    newRow[9] = new Date();
    updates.push({ rowNumber: idx + 2, values: newRow });
  });
  updates.forEach(update => {
    sheet.getRange(update.rowNumber, 1, 1, update.values.length).setValues([update.values]);
  });
  return { billingMonth, updated: updates.length };
}

const BILLING_PAYMENT_PDF_STATUS_LABELS = {
  '回収済み': 'OK',
  '預金口座振替依頼書なし': 'NO_DOCUMENT',
  '資金不足': 'INSUFFICIENT',
  '取引なし': 'NOT_FOUND'
};

function parseBillingPaymentResultPdf(pdfBlob) {
  const content = pdfBlob.getDataAsString();
  const lines = content.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  const entries = [];
  lines.forEach(line => {
    const paidMatch = line.match(/([0-9,]+)円/);
    const statusLabel = Object.keys(BILLING_PAYMENT_PDF_STATUS_LABELS).find(label => line.indexOf(label) >= 0) || '';
    const nameMatch = line.match(/^(.*?)(回収済み|預金口座振替依頼書なし|資金不足|取引なし)/);
    if (!nameMatch) return;
    const name = normalizeBillingNameKey_(nameMatch[1]);
    const paidAmount = paidMatch ? Number(paidMatch[1].replace(/,/g, '')) : 0;
    const bankStatus = BILLING_PAYMENT_PDF_STATUS_LABELS[statusLabel] || '';
    entries.push({ nameKanji: name, paidAmount, bankStatus, statusLabel });
  });
  return entries;
}

function applyPaymentResultPdf(billingMonth, pdfBlob, billingJson) {
  const parsed = parseBillingPaymentResultPdf(pdfBlob);
  const nameIndex = {};
  (billingJson || []).forEach(item => {
    const key = normalizeBillingNameKey_(item.nameKanji);
    if (key && !nameIndex[key]) {
      nameIndex[key] = item;
    }
  });

  const matched = [];
  parsed.forEach(entry => {
    const key = normalizeBillingNameKey_(entry.nameKanji);
    const target = nameIndex[key];
    if (!target) return;
    matched.push({
      patientId: target.patientId,
      billingMonth,
      paidAmount: entry.paidAmount,
      unpaidAmount: normalizeBillingAmount_(target) - entry.paidAmount,
      bankStatus: entry.bankStatus,
      statusLabel: entry.statusLabel
    });
  });

  const statusMap = matched.reduce((map, entry) => {
    map[entry.patientId] = entry;
    return map;
  }, {});
  const historyResult = applyPaymentResultsToHistory(billingMonth, statusMap);

  return {
    billingMonth,
    parsedCount: parsed.length,
    matched: matched.length,
    updated: historyResult.updated,
    entries: matched
  };
}
