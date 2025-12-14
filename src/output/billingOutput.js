/***** Output layer: billing invoice PDF generation *****/

const INVOICE_PARENT_FOLDER_ID = '1EG-GB3PbaUr9C1LJWlaf_idqoYF-19Ux';
const INVOICE_FILE_PREFIX = '請求書';
const TRANSPORT_PRICE = (typeof BILLING_TRANSPORT_UNIT_PRICE !== 'undefined')
  ? BILLING_TRANSPORT_UNIT_PRICE
  : 33;
const INVOICE_TREATMENT_UNIT_PRICE_BY_BURDEN = { 1: 417, 2: 834, 3: 1251 };
const INVOICE_UNIT_PRICE_FALLBACK = (typeof BILLING_UNIT_PRICE !== 'undefined') ? BILLING_UNIT_PRICE : 4170;

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
    if (String(burdenRate).trim() === '自費') return '自費';
    const num = Number(burdenRate);
    if (Number.isFinite(num)) {
      if (num > 0 && num < 1) return Math.round(num * 10);
    if (num >= 1 && num < 10) return Math.round(num);
    if (num >= 10 && num <= 100) return Math.round(num / 10);
  }
  const normalized = String(burdenRate).normalize('NFKC').replace(/\s+/g, '').replace('％', '%');
  const hasPercent = normalized.indexOf('%') >= 0;
  const numericText = normalized.replace(/[^0-9.]/g, '');
  const parsed = Number(numericText);
  if (!Number.isFinite(parsed)) return 0;
  if (parsed === 0) return 0;
  if (hasPercent) return Math.round(parsed / 10);
  if (parsed > 0 && parsed < 10) return Math.round(parsed);
  if (parsed >= 10 && parsed <= 100) return Math.round(parsed / 10);
  return 0;
  };

const resolveInvoiceUnitPriceForOutput_ = typeof resolveInvoiceUnitPrice_ === 'function'
  ? resolveInvoiceUnitPrice_
  : function fallbackResolveInvoiceUnitPriceForOutput_(insuranceType, burdenRate, manualUnitPrice, medicalAssistance, patientUnitPrice) {
    const type = String(insuranceType || '').trim();
    if (type === 'マッサージ') return 0;
    const normalizedManual = normalizeInvoiceMoney_(manualUnitPrice);
    const hasManual = Number.isFinite(normalizedManual) && normalizedManual !== 0;
    if (hasManual) return normalizedManual;
    const assistance = normalizeInvoiceMedicalAssistanceFlag_(medicalAssistance);
    if (assistance === 1) return 0;
    const isLifeProtection = ['生保', '生活保護', '生活扶助'].indexOf(type) >= 0;
    if (isLifeProtection) return 0;
    const normalizedBurdenRate = normalizeInvoiceBurdenRateInt_(burdenRate);
    const isSelfPaid = type === '自費' || normalizedBurdenRate === '自費';
    if (isSelfPaid) return 0;
    const normalizedPatientPrice = normalizeInvoiceMoney_(patientUnitPrice);
    if (Number.isFinite(normalizedPatientPrice) && normalizedPatientPrice !== 0) return normalizedPatientPrice;
    return INVOICE_UNIT_PRICE_FALLBACK;
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

function normalizeKana_(value) {
  return String(value || '')
    .normalize('NFKC')
    .trim();
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
  if (value === 1 || value === '1' || value === true) return 1;
  return 0;
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
    { label: '交通費', detail: breakdown.transportDetail || (formatBillingCurrency_(TRANSPORT_PRICE) + '円 × ' + visits + '回'), amount: breakdown.transportAmount }
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
  const normalizedMedicalAssistance = normalizeInvoiceMedicalAssistanceFlag_(params && params.medicalAssistance);
  const carryOverAmount = normalizeBillingCarryOver_(params);
  const manualUnitPrice = params && params.hasOwnProperty('manualUnitPrice')
    ? params.manualUnitPrice
    : params && params.unitPrice;
  const hasManualTransportInput = params && Object.prototype.hasOwnProperty.call(params, 'manualTransportAmount');
  const manualTransportInput = hasManualTransportInput ? params.manualTransportAmount : null;
  const manualTransportAmount = (manualTransportInput === ''
    || manualTransportInput === null
    || manualTransportInput === undefined
    || !hasManualTransportInput)
    ? null
    : normalizeInvoiceMoney_(manualTransportInput);
  const patientUnitPrice = params && params.unitPrice;
  const treatmentUnitPrice = resolveInvoiceUnitPriceForOutput_(
    insuranceType,
    burdenRateInt,
    manualUnitPrice,
    normalizedMedicalAssistance,
    patientUnitPrice || INVOICE_TREATMENT_UNIT_PRICE_BY_BURDEN[burdenRateInt]
  );

  const hasChargeableUnitPrice = Number.isFinite(treatmentUnitPrice) && treatmentUnitPrice !== 0;
  const treatmentAmountFull = visits > 0 && hasChargeableUnitPrice ? treatmentUnitPrice * visits : 0;
  const isSelfPaid = insuranceType === '自費' || burdenRateInt === '自費';
  const defaultBurdenUnitPrice = INVOICE_TREATMENT_UNIT_PRICE_BY_BURDEN[burdenRateInt];
  const usesBurdenAdjustedUnitPrice = Number.isFinite(defaultBurdenUnitPrice)
    && treatmentUnitPrice === defaultBurdenUnitPrice;
  const burdenMultiplier = isSelfPaid || usesBurdenAdjustedUnitPrice
    ? 1
    : (typeof normalizeBurdenMultiplier_ === 'function'
      ? normalizeBurdenMultiplier_(burdenRateInt, insuranceType)
      : (insuranceType === '自費' ? 1 : (burdenRateInt > 0 ? burdenRateInt / 10 : 0)));
  const treatmentAmount = isSelfPaid
    ? treatmentAmountFull
    : roundToNearestTen_(treatmentAmountFull * burdenMultiplier);
  const transportAmount = (manualTransportInput !== '' && manualTransportInput !== null && manualTransportInput !== undefined
    && Number.isFinite(manualTransportAmount))
    ? manualTransportAmount
    : visits > 0 && hasChargeableUnitPrice ? TRANSPORT_PRICE * visits : 0;
  const transportDetail = (manualTransportInput !== '' && manualTransportInput !== null && manualTransportInput !== undefined
    && hasManualTransportInput)
    ? '手動入力'
    : formatBillingCurrency_(TRANSPORT_PRICE) + '円 × ' + visits + '回';
  const selfPayItems = Array.isArray(params && params.selfPayItems)
    ? params.selfPayItems
    : (params && params.manualSelfPayAmount ? [{ type: '自費', amount: params.manualSelfPayAmount }] : []);
  const selfPayTotal = selfPayItems.reduce((sum, entry) => sum + (normalizeInvoiceMoney_(entry.amount) || 0), 0);
  const grandTotal = carryOverAmount + treatmentAmount + transportAmount + selfPayTotal;

  return { treatmentUnitPrice, treatmentAmount, transportAmount, transportDetail, grandTotal, visits, selfPayItems, selfPayTotal };
}

  function buildBillingInvoiceHtml_(item, billingMonth) {
    const targetMonth = billingMonth || (item && item.billingMonth) || '';
    const breakdown = calculateInvoiceChargeBreakdown_(Object.assign({}, item, { billingMonth: targetMonth }));
    const monthLabel = normalizeBillingMonthLabel_(targetMonth);
    const visits = breakdown.visits || 0;
  const treatmentUnitPrice = breakdown.treatmentUnitPrice || 0;
  const transportUnitPrice = TRANSPORT_PRICE;
  const carryOverAmount = normalizeBillingCarryOver_(item);
  const transportDetail = breakdown.transportDetail || (formatBillingCurrency_(transportUnitPrice) + '円 × ' + visits + '回');
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
    `<p>交通費（${transportDetail}）: ${formatBillingCurrency_(breakdown.transportAmount)}円</p>`,
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

function buildBankTransferRowsForBilling_(billingJson, bankInfoByName, patientMap, billingMonth, bankStatuses) {
  const rows = [];
  let skipped = 0;
  let passed = 0;
  let total = 0;
  const skipReasons = {
    invalidBankCode: 0,
    invalidBranchCode: 0,
    invalidAccountNumber: 0
  };
  const billingMonthKey = billingMonth || (billingJson && billingJson.length ? billingJson[0].billingMonth : '');

  (billingJson || []).forEach(item => {
    const pid = item && item.patientId ? String(item.patientId).trim() : '';
    if (!pid) return;
    total += 1;
    const patient = patientMap && patientMap[pid] ? patientMap[pid] : {};
    const nameKanji = item && item.nameKanji ? String(item.nameKanji).trim() : '';
    const nameKey = normalizeBillingNameKey_(nameKanji);
    const bankLookup = bankInfoByName && nameKey ? bankInfoByName[nameKey] : null;

    const pickWithPriority = (resolver, fallbackValue) => {
      const sources = [bankLookup, patient, item];
      for (let i = 0; i < sources.length; i += 1) {
        const source = sources[i];
        if (!source) continue;
        const value = typeof resolver === 'function' ? resolver(source) : source[resolver];
        if (value !== undefined && value !== null && String(value).trim() !== '') {
          return value;
        }
      }
      return fallbackValue;
    };

    const rawBankCode = pickWithPriority('bankCode', '');
    const bankCode = String(rawBankCode).replace(/\D/g, '').padStart(4, '0');
    const rawBranchCode = pickWithPriority('branchCode', '');
    const branchCode = String(rawBranchCode).replace(/\D/g, '').padStart(3, '0');
    const regulationCode = pickWithPriority('regulationCode', 1);
    const mergedNameKanji = pickWithPriority('nameKanji', nameKanji);
    const rawNameKana = pickWithPriority('nameKana', '');
    const nameKana = rawNameKana ? normalizeKana_(rawNameKana) : normalizeKana_(mergedNameKanji);
    const rawAccountNumber = pickWithPriority('accountNumber', '');
    const accountNumber = String(rawAccountNumber).replace(/\D/g, '').padStart(7, '0');
    const isNew = normalizeZeroOneFlag_(pickWithPriority('isNew', ''));
    const statusEntry = bankStatuses && pid ? bankStatuses[pid] : null;
    const paidStatus = item && item.paidStatus ? item.paidStatus : (statusEntry && statusEntry.paidStatus ? statusEntry.paidStatus : '');

    const bankCodeInvalid = bankCode.length !== 4;
    const branchCodeInvalid = branchCode.length !== 3;
    const accountNumberInvalid = accountNumber.length !== 7;

    if (bankCodeInvalid || branchCodeInvalid || accountNumberInvalid) {
      skipped += 1;
      if (bankCodeInvalid) skipReasons.invalidBankCode += 1;
      if (branchCodeInvalid) skipReasons.invalidBranchCode += 1;
      if (accountNumberInvalid) skipReasons.invalidAccountNumber += 1;
      return;
    }

    passed += 1;
    rows.push({
      billingMonth: billingMonthKey,
      patientId: pid,
      nameKanji: mergedNameKanji,
      bankCode,
      branchCode,
      regulationCode,
      accountNumber,
      nameKana,
      isNew,
      paidStatus
    });
  });

  return {
    billingMonth: billingMonthKey,
    rows,
    skipped,
    total,
    passed,
    skipReasons
  };
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
  ensureColumn('領収状態', ['領収状態', '領収', 'paidStatus']);

  return { columns: resolved, headers: workingHeaders };
}

function exportBankTransferRows_(billingMonth, rowObjects, bankStatuses) {
  const ensured = ensureBankTransferSheet_();
  const sheet = ensured.sheet;
  const { columns, headers } = resolveBankTransferColumns_(sheet, ensured.headers);
  const colCount = Math.max(sheet.getLastColumn(), headers.length, Math.max.apply(null, Object.values(columns)));

  const lastRow = sheet.getLastRow();
  const existingValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, colCount).getValues() : [];
  const keyForRow = (row) => {
    const month = columns['請求月'] ? String(row[columns['請求月'] - 1] || '').trim() : '';
    const pid = columns['番号'] ? String(row[columns['番号'] - 1] || '').trim() : '';
    return month && pid ? `${month}::${pid}` : '';
  };

  const workingRowsByKey = existingValues.reduce((map, row) => {
    const key = keyForRow(row);
    if (key && !map.has(key)) map.set(key, row);
    return map;
  }, new Map());

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
    if (columns['領収状態']) {
      const statusEntry = bankStatuses && obj.patientId ? bankStatuses[obj.patientId] : null;
      const existingPaid = row[columns['領収状態'] - 1] || '';
      const paidStatus = (obj.paidStatus != null && obj.paidStatus !== '')
        ? obj.paidStatus
        : (statusEntry && statusEntry.paidStatus ? statusEntry.paidStatus : existingPaid);
      row[columns['領収状態'] - 1] = paidStatus || '';
    }
    return row;
  });

  mapped.forEach(row => {
    const key = keyForRow(row);
    if (key) {
      const existingRow = workingRowsByKey.get(key);
      const mergedRow = existingRow ? existingRow.slice() : new Array(colCount).fill('');
      ['請求月', '番号', '氏名（漢字）', '銀行コード', '支店コード', '規定コード', '口座番号', '氏名（カナ）', '新規フラグ']
        .forEach(label => {
          if (columns[label]) {
            mergedRow[columns[label] - 1] = row[columns[label] - 1];
          }
        });
      const existingPaidStatus = columns['領収状態'] ? mergedRow[columns['領収状態'] - 1] : '';
      if (columns['領収状態']) {
        const statusEntry = bankStatuses && row[columns['番号'] - 1] ? bankStatuses[row[columns['番号'] - 1]] : null;
        const resolvedPaidStatus = (row[columns['領収状態'] - 1] != null && row[columns['領収状態'] - 1] !== '')
          ? row[columns['領収状態'] - 1]
          : (statusEntry && statusEntry.paidStatus ? statusEntry.paidStatus : existingPaidStatus || '');
        mergedRow[columns['領収状態'] - 1] = resolvedPaidStatus;
      }
      workingRowsByKey.set(key, mergedRow);
    }
  });

  const sortedKeys = Array.from(workingRowsByKey.keys()).sort();
  const workingRows = sortedKeys.map(key => workingRowsByKey.get(key));
  const dataRowCount = Math.max(0, lastRow - 1);
  const maxRowCount = Math.max(dataRowCount, workingRows.length);

  if (maxRowCount > 0) {
    sheet.getRange(2, 1, maxRowCount, colCount).clearContent();
  }
  if (workingRows.length) {
    sheet.getRange(2, 1, workingRows.length, colCount).setValues(workingRows);
  }

  return { billingMonth, inserted: mapped.length };
}

  function logPreparedBankPayloadStatus_(prepared) {
    const requiredKeys = ['billingJson', 'visitsByPatient', 'totalsByPatient', 'carryOverByPatient', 'unpaidHistory', 'bankAccountInfoByPatient'];
    const normalized = normalizePreparedBilling_(prepared) || {};
    const missing = requiredKeys.filter(key => {
      const value = normalized[key];
      if (key === 'billingJson') {
        return !Array.isArray(value) || value.length === 0;
      }
      if (Array.isArray(value)) return value.length === 0;
      return !value || (typeof value === 'object' && Object.keys(value).length === 0);
    });
    if (missing.length) {
      billingLogger_.log('[billing] Prepared payload is incomplete (missing: ' + missing.join(', ') + ')');
    }
    if (normalized && normalized.carryOverLedgerMeta && (normalized.carryOverLedgerMeta.wasAutoCreated || normalized.carryOverLedgerMeta.headerInserted)) {
      billingLogger_.log('[billing] CarryOverLedger sheet missing → using fallback model');
    }
  }

  function exportBankTransferDataForPrepared_(prepared) {
    const normalized = normalizePreparedBilling_(prepared);
    if (!normalized) {
      throw new Error('銀行データを生成できません。請求データが未生成です。先に「請求データを集計」を実行してください。');
    }

    if (!Array.isArray(normalized.billingJson)) {
      throw new Error('銀行データを生成できません。請求データの形式が不正です。先に「請求データを集計」を実行してください。');
    }
    logPreparedBankPayloadStatus_(normalized);

    if (normalized.billingJson.length === 0) {
      billingLogger_.log('[billing] exportBankTransferDataForPrepared_: billingJson empty for ' + (normalized.billingMonth || ''));
      return { billingMonth: normalized.billingMonth || '', rows: [], inserted: 0, skipped: 0, message: '当月の請求対象はありません' };
    }

    let bankInfoByName = normalized.bankInfoByName || {};
    let patientMap = normalized.patients || normalized.patientMap || {};
    let bankStatuses = normalized.bankStatuses || {};

    if (!Object.keys(bankInfoByName).length || !Object.keys(patientMap).length) {
      const source = getBillingSourceData(normalized.billingMonth);
      bankInfoByName = source.bankInfoByName || bankInfoByName;
      patientMap = source.patients || source.patientMap || patientMap;
      bankStatuses = source.bankStatuses || bankStatuses;
      if (!normalized.billingJson || !normalized.billingJson.length) {
        normalized.billingJson = generateBillingJsonFromSource(source);
      }
    }

    const buildResult = buildBankTransferRowsForBilling_(normalized.billingJson, bankInfoByName, patientMap, normalized.billingMonth, bankStatuses);
    const outputResult = exportBankTransferRows_(buildResult.billingMonth, buildResult.rows, bankStatuses);
    const combined = Object.assign({}, buildResult, outputResult);

    if (!buildResult.rows.length) {
      const reasonSummary = buildResult.skipReasons || {};
      const parts = [
        '銀行CSVが生成されませんでした',
        `総件数: ${buildResult.total || 0}`,
        `有効: ${buildResult.passed || 0}`,
        `銀行コード不正: ${reasonSummary.invalidBankCode || 0}`,
        `支店コード不正: ${reasonSummary.invalidBranchCode || 0}`,
        `口座番号不正: ${reasonSummary.invalidAccountNumber || 0}`
      ];
      combined.message = parts.join(' / ');
    }

    return combined;
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
