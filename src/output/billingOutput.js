/***** Output layer: billing invoice PDF generation *****/

const INVOICE_PARENT_FOLDER_ID = '1EG-GB3PbaUr9C1LJWlaf_idqoYF-19Ux';
const INVOICE_FILE_PREFIX = '請求書';
const BILLING_TRANSPORT_UNIT_PRICE = (typeof globalThis !== 'undefined' && typeof globalThis.BILLING_TRANSPORT_UNIT_PRICE !== 'undefined')
  ? globalThis.BILLING_TRANSPORT_UNIT_PRICE
  : 33;

function convertSpreadsheetToExcelBlob_(file, exportName) {
  if (!file || typeof file.getMimeType !== 'function' || file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    throw new Error('スプレッドシート以外のファイルをExcelに変換することはできません');
  }

  const blob = file.getBlob();
  const name = (exportName && String(exportName).trim()) || 'export';
  return blob.getAs(MimeType.MICROSOFT_EXCEL).setName(name + '.xlsx');
}

function normalizeBillingAmount_(item) {
  if (!item) return 0;
  if (item.grandTotal != null && item.grandTotal !== '') return Number(item.grandTotal) || 0;
  if (item.treatmentAmount != null && item.treatmentAmount !== '') return Number(item.treatmentAmount) || 0;
  return 0;
}

function normalizeBillingNameKey_(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}

function normalizeInvoiceMoney_(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }
  const text = String(value || '').replace(/,/g, '').trim();
  if (!text) return 0;
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
}

function normalizeInvoiceVisitCount_(value) {
  const num = Number(value && value.visitCount != null ? value.visitCount : value);
  return Number.isFinite(num) && num > 0 ? num : 0;
}

function normalizeBillingCarryOver_(item) {
  if (!item) return 0;
  if (item.carryOverAmount != null && item.carryOverAmount !== '') return Number(item.carryOverAmount) || 0;
  if (item.raw && item.raw.carryOverAmount != null) return Number(item.raw.carryOverAmount) || 0;
  return 0;
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

function formatInvoiceFileName_(billingMonth) {
  const digits = (billingMonth ? String(billingMonth) : '').replace(/\D/g, '');
  if (digits.length >= 6) {
    const year = digits.slice(0, 4);
    const month = digits.slice(4, 6).padStart(2, '0');
    return year + '-' + month + '_' + INVOICE_FILE_PREFIX + '.pdf';
  }
  return INVOICE_FILE_PREFIX + '.pdf';
}

function buildInvoiceTemplateData_(item) {
  const visits = normalizeInvoiceVisitCount_(item && item.visitCount);
  const unitPrice = normalizeInvoiceMoney_(item && item.unitPrice);
  const monthLabel = normalizeBillingMonthLabel_(item && item.billingMonth);
  const rows = [
    { label: '前月繰越', detail: '', amount: normalizeBillingCarryOver_(item) },
    { label: '施術料', detail: formatBillingCurrency_(unitPrice) + '円 × ' + visits + '回', amount: normalizeInvoiceMoney_(item && item.treatmentAmount) },
    { label: '交通費', detail: formatBillingCurrency_(BILLING_TRANSPORT_UNIT_PRICE) + '円 × ' + visits + '回', amount: normalizeInvoiceMoney_(item && item.transportAmount) }
  ];

  return Object.assign({}, item, {
    monthLabel,
    rows,
    grandTotal: normalizeInvoiceMoney_(item && item.grandTotal)
  });
}

function createInvoicePdfBlob_(item) {
  const template = HtmlService.createTemplateFromFile('invoice_template');
  template.data = buildInvoiceTemplateData_(item || {});
  const html = template.evaluate().setWidth(1240).setHeight(1754);
  const fileName = formatInvoiceFileName_(item && item.billingMonth);
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

function ensureInvoiceFolderForPatient_(patientName, billingMonth) {
  const root = ensureInvoiceRootFolder_();
  const safeName = (patientName || '患者未設定').replace(/[\\/]/g, '_');
  const digits = (billingMonth ? String(billingMonth) : '').replace(/\D/g, '');
  const year = (digits.length >= 6 ? digits.slice(0, 4) : '未設定年').replace(/[\\/]/g, '_');
  const patientFolder = ensureSubFolder_(root, safeName);
  return ensureSubFolder_(patientFolder, year);
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
  const folder = ensureInvoiceFolderForPatient_(item && item.nameKanji, item && item.billingMonth);
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
