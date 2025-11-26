/***** Output layer: billing Excel/CSV/history generation *****/

const BILLING_EXCEL_COLUMNS = {
  nameKanji: columnLetterToNumber_('F'),
  bankCode: columnLetterToNumber_('N'),
  branchCode: columnLetterToNumber_('O'),
  constantOne: columnLetterToNumber_('P'),
  accountNumber: columnLetterToNumber_('Q'),
  nameKana: columnLetterToNumber_('R'),
  billingAmount: columnLetterToNumber_('S'),
  isNew: columnLetterToNumber_('U')
};

function normalizeBillingAmount_(item) {
  if (!item) return 0;
  if (item.grandTotal != null && item.grandTotal !== '') return item.grandTotal;
  if (item.billingAmount != null && item.billingAmount !== '') return item.billingAmount;
  return 0;
}

function normalizeBillingNameKey_(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}

function buildBillingExcelRows_(billingJson) {
  if (!Array.isArray(billingJson)) {
    throw new Error('請求データが不正です');
  }
  const maxCol = BILLING_EXCEL_COLUMNS.isNew;
  const header = Array(maxCol).fill('');
  header[BILLING_EXCEL_COLUMNS.nameKanji - 1] = 'nameKanji';
  header[BILLING_EXCEL_COLUMNS.bankCode - 1] = 'bankCode';
  header[BILLING_EXCEL_COLUMNS.branchCode - 1] = 'branchCode';
  header[BILLING_EXCEL_COLUMNS.constantOne - 1] = '1固定';
  header[BILLING_EXCEL_COLUMNS.accountNumber - 1] = 'accountNumber';
  header[BILLING_EXCEL_COLUMNS.nameKana - 1] = 'nameKana';
  header[BILLING_EXCEL_COLUMNS.billingAmount - 1] = 'billingAmount/grandTotal';
  header[BILLING_EXCEL_COLUMNS.isNew - 1] = 'isNew';

  const sorted = billingJson.slice().sort((a, b) => {
    const aid = Number(a && a.patientId);
    const bid = Number(b && b.patientId);
    if (isNaN(aid) && isNaN(bid)) return 0;
    if (isNaN(aid)) return 1;
    if (isNaN(bid)) return -1;
    return aid - bid;
  });

  const rows = sorted.map(item => {
    const row = Array(maxCol).fill('');
    row[BILLING_EXCEL_COLUMNS.nameKanji - 1] = item && item.nameKanji ? item.nameKanji : '';
    row[BILLING_EXCEL_COLUMNS.bankCode - 1] = item && item.bankCode ? item.bankCode : '';
    row[BILLING_EXCEL_COLUMNS.branchCode - 1] = item && item.branchCode ? item.branchCode : '';
    row[BILLING_EXCEL_COLUMNS.constantOne - 1] = 1;
    row[BILLING_EXCEL_COLUMNS.accountNumber - 1] = item && item.accountNumber ? item.accountNumber : '';
    row[BILLING_EXCEL_COLUMNS.nameKana - 1] = item && item.nameKana ? item.nameKana : '';
    row[BILLING_EXCEL_COLUMNS.billingAmount - 1] = normalizeBillingAmount_(item);
    row[BILLING_EXCEL_COLUMNS.isNew - 1] = item && item.isNew ? 1 : 0;
    return row;
  });

  return [header].concat(rows);
}

function copyTemplateSheet_(templateSheetName, destinationSpreadsheet) {
  const templateName = templateSheetName || '請求一覧_TEMPLATE';
  const workbook = ss();
  const templateSheet = workbook.getSheetByName(templateName);
  if (!templateSheet) {
    throw new Error('テンプレートシートが見つかりません: ' + templateName);
  }

  const targetSpreadsheet = destinationSpreadsheet || SpreadsheetApp.create('請求一覧_出力中');
  const copiedSheet = templateSheet.copyTo(targetSpreadsheet);
  const newName = '請求一覧_出力中_' + Utilities.getUuid().slice(0, 8);
  copiedSheet.setName(newName);
  targetSpreadsheet.setActiveSheet(copiedSheet);
  targetSpreadsheet.moveActiveSheet(targetSpreadsheet.getSheets().length);

  return { sheet: copiedSheet, spreadsheet: targetSpreadsheet };
}

function writeBillingExcelRows_(sheet, rows) {
  if (!sheet || !Array.isArray(rows) || !rows.length) {
    return 0;
  }
  const startRow = 2; // 1行目はテンプレ側ヘッダ
  const startCol = 1;
  sheet.getRange(startRow, startCol, rows.length, rows[0].length).setValues(rows);
  return rows.length;
}

function createBillingExcelFile(billingJson, options) {
  const opts = options || {};
  const billingMonth = opts.billingMonth || (Array.isArray(billingJson) && billingJson.length && billingJson[0].billingMonth) || '';
  const baseName = opts.fileName || (billingMonth ? '請求一覧_' + billingMonth : '請求一覧');
  const rows = buildBillingExcelRows_(billingJson);
  const valueRows = rows.length > 1 ? rows.slice(1) : [];
  const templateSheetName = opts.templateSheetName || '請求一覧_TEMPLATE';
  const outputSheetName = opts.outputSheetName || '請求一覧';

  const { sheet, spreadsheet } = copyTemplateSheet_(templateSheetName, SpreadsheetApp.create(baseName));
  spreadsheet.getSheets().forEach(s => {
    if (s.getSheetId() !== sheet.getSheetId()) {
      spreadsheet.deleteSheet(s);
    }
  });
  sheet.setName(outputSheetName);
  if (valueRows.length) {
    writeBillingExcelRows_(sheet, valueRows);
  }

  const tempFile = DriveApp.getFileById(spreadsheet.getId());
  let folder = null;
  try {
    folder = getParentFolder_();
  } catch (e) {
    folder = null;
  }
  if (folder) {
    folder.addFile(tempFile);
    DriveApp.getRootFolder().removeFile(tempFile);
  }

  const excelBlob = tempFile.getBlob().getAs(MimeType.MICROSOFT_EXCEL).setName(baseName + '.xlsx');
  const outFile = folder ? folder.createFile(excelBlob) : DriveApp.createFile(excelBlob);
  tempFile.setTrashed(true);

  return {
    fileId: outFile.getId(),
    url: outFile.getUrl(),
    name: outFile.getName(),
    rowCount: rows.length ? rows.length - 1 : 0,
    billingMonth
  };
}

function escapeBillingCsvCell_(value) {
  const text = value == null ? '' : String(value);
  if (/[",\n]/.test(text)) {
    return '"' + text.replace(/"/g, '""') + '"';
  }
  return text;
}

function createBillingCsvFile(billingJson, options) {
  const opts = options || {};
  const billingMonth = opts.billingMonth || (Array.isArray(billingJson) && billingJson.length && billingJson[0].billingMonth) || '';
  const baseName = opts.fileName || (billingMonth ? '口座振替_' + billingMonth : '口座振替');
  const sorted = Array.isArray(billingJson) ? billingJson.slice().sort((a, b) => {
    const aid = Number(a && a.patientId);
    const bid = Number(b && b.patientId);
    if (isNaN(aid) && isNaN(bid)) return 0;
    if (isNaN(aid)) return 1;
    if (isNaN(bid)) return -1;
    return aid - bid;
  }) : [];
  const rows = sorted.map(item => [
    item && item.nameKanji ? item.nameKanji : '',
    item && item.nameKana ? item.nameKana : '',
    normalizeBillingAmount_(item)
  ]);
  const csv = rows.map(row => row.map(escapeBillingCsvCell_).join(',')).join('\r\n');
  const blob = Utilities.newBlob(csv, 'text/csv', baseName + '.csv');
  let folder = null;
  try {
    folder = getParentFolder_();
  } catch (e) {
    folder = null;
  }
  const file = folder ? folder.createFile(blob) : DriveApp.createFile(blob);
  return {
    fileId: file.getId(),
    url: file.getUrl(),
    name: file.getName(),
    rowCount: rows.length,
    billingMonth
  };
}

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
    const billingAmount = item && item.billingAmount != null ? item.billingAmount : 0;
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
  if (!pdfBlob) {
    throw new Error('入金結果PDFが指定されていません');
  }
  const text = pdfBlob.getDataAsString('Shift_JIS');
  const normalized = text.replace(/\r\n/g, '\n').replace(/[\t ]+/g, ' ');
  const statusKeys = Object.keys(BILLING_PAYMENT_PDF_STATUS_LABELS).join('|');
  const regex = new RegExp('([\\p{sc=Han}\\p{L}\\p{N}\\s・･()（）]+?)\\s+([0-9,]+)\\s+(' + statusKeys + ')', 'gu');
  const entries = [];
  let match = null;
  while ((match = regex.exec(normalized)) !== null) {
    const name = normalizeBillingNameKey_(match[1]);
    if (!name) continue;
    const paidAmount = Number(String(match[2] || '0').replace(/,/g, '')) || 0;
    const statusLabel = match[3];
    const bankStatus = BILLING_PAYMENT_PDF_STATUS_LABELS[statusLabel] || '';
    entries.push({ nameKanji: name, paidAmount, bankStatus, statusLabel });
  }

  if (!entries.length) {
    const lines = normalized.split(/\n+/).map(line => line.trim()).filter(Boolean);
    lines.forEach(line => {
      Object.keys(BILLING_PAYMENT_PDF_STATUS_LABELS).forEach(label => {
        if (line.indexOf(label) < 0) return;
        const paidMatch = line.match(/([0-9][0-9,]*)/);
        const paidAmount = paidMatch ? Number(paidMatch[1].replace(/,/g, '')) || 0 : 0;
        const name = normalizeBillingNameKey_(line.replace(label, '').replace(paidMatch ? paidMatch[1] : '', ''));
        if (!name) return;
        entries.push({ nameKanji: name, paidAmount, bankStatus: BILLING_PAYMENT_PDF_STATUS_LABELS[label], statusLabel: label });
      });
    });
  }

  return entries;
}

function buildPaymentResultStatusMap_(parsedEntries, billingJson) {
  const nameIndex = {};
  (billingJson || []).forEach(item => {
    if (item && item.nameKanji) {
      nameIndex[normalizeBillingNameKey_(item.nameKanji)] = item;
    }
  });

  const statusMap = {};
  const unmatched = [];

  parsedEntries.forEach(entry => {
    const key = normalizeBillingNameKey_(entry.nameKanji);
    const target = nameIndex[key];
    if (!target || !target.patientId) {
      unmatched.push(entry);
      return;
    }
    const paidAmount = Number(entry.paidAmount) || 0;
    const grandTotal = normalizeBillingAmount_(target);
    statusMap[target.patientId] = {
      bankStatus: entry.bankStatus || '',
      paidAmount,
      unpaidAmount: grandTotal - paidAmount
    };
  });

  return { statusMap, unmatched, matched: Object.keys(statusMap).length };
}

function applyPaymentResultPdf(billingMonth, pdfBlob, billingJson) {
  const parsedEntries = parseBillingPaymentResultPdf(pdfBlob);
  const mapping = buildPaymentResultStatusMap_(parsedEntries, billingJson);
  const historyResult = applyPaymentResultsToHistory(billingMonth, mapping.statusMap);
  return Object.assign({ billingMonth }, historyResult, {
    parsedCount: parsedEntries.length,
    matched: mapping.matched,
    unmatched: mapping.unmatched
  });
}

const BILLING_COMBINED_INVOICE_TEMPLATE_ID = '1CLy0facEX8_CFFiswIywSGxvan0dUKVw';
const BILLING_STANDARD_INVOICE_TEMPLATE_ID = '15__-XmSsDcMV2mFzsekxuGMYcNWtm2CN';

const BILLING_PDF_OVERLAY_POSITIONS = {
  nameKanji: { x: 0.63, y: 0.26, fontSize: 16, align: 'left', weight: '700' },
  address: { x: 0.63, y: 0.23, fontSize: 11, align: 'left' },
  billingMonth: { x: 0.22, y: 0.35, fontSize: 13, align: 'left' },
  combinedTotal: { x: 0.80, y: 0.63, fontSize: 16, align: 'right', weight: '700' },
  carryOver: { x: 0.80, y: 0.69, fontSize: 14, align: 'right' },
  combineNote: { x: 0.63, y: 0.75, fontSize: 12, align: 'left', weight: '600' }
};

function resolveBillingPdfTemplates_(options) {
  const props = PropertiesService.getScriptProperties();
  return {
    combined: (options && options.combinedTemplateId) || props.getProperty('BILLING_COMBINED_INVOICE_TEMPLATE_ID') ||
      BILLING_COMBINED_INVOICE_TEMPLATE_ID,
    standard: (options && options.standardTemplateId) || props.getProperty('BILLING_STANDARD_INVOICE_TEMPLATE_ID') ||
      BILLING_STANDARD_INVOICE_TEMPLATE_ID
  };
}

function normalizeBillingCarryOver_(item) {
  if (!item) return 0;
  if (item.carryOverAmount != null && item.carryOverAmount !== '') return Number(item.carryOverAmount) || 0;
  if (item.raw && item.raw.carryOverAmount != null) return Number(item.raw.carryOverAmount) || 0;
  return 0;
}

function resolveBillingAddress_(item) {
  if (!item || !item.raw) return '';
  const raw = item.raw;
  const candidates = ['住所', '住所1', '住所２', '住所2', 'address', 'Address'];
  for (let i = 0; i < candidates.length; i++) {
    const key = candidates[i];
    if (raw.hasOwnProperty(key) && raw[key] != null && String(raw[key]).trim()) {
      return String(raw[key]).trim();
    }
  }
  return '';
}

function formatBillingCurrency_(value) {
  const num = Number(value);
  if (!isFinite(num)) return '0';
  return Math.round(num).toLocaleString('ja-JP');
}

function createBillingPdfOverlay_(pngBlob, overlay) {
  const image = ImagesService.openImage(pngBlob);
  const width = image.getWidth();
  const height = image.getHeight();
  const backgroundUrl = 'data:image/png;base64,' + Utilities.base64Encode(pngBlob.getBytes());

  const pos = BILLING_PDF_OVERLAY_POSITIONS;
  const fields = [];
  if (overlay.address) {
    fields.push(Object.assign({ text: overlay.address }, pos.address));
  }
  if (overlay.nameKanji) {
    fields.push(Object.assign({ text: overlay.nameKanji }, pos.nameKanji));
  }
  if (overlay.billingMonth) {
    fields.push(Object.assign({ text: overlay.billingMonth }, pos.billingMonth));
  }
  if (overlay.combinedTotal) {
    fields.push(Object.assign({ text: overlay.combinedTotal }, pos.combinedTotal));
  }
  if (overlay.carryOver) {
    fields.push(Object.assign({ text: overlay.carryOver }, pos.carryOver));
  }
  if (overlay.combineNote) {
    fields.push(Object.assign({ text: overlay.combineNote }, pos.combineNote));
  }

  const textLayer = fields.map(field => {
    const left = Math.round(field.x * width);
    const top = Math.round(field.y * height);
    const fontSize = field.fontSize || 12;
    const weight = field.weight || '400';
    const align = field.align || 'left';
    const content = escapeHtml_ ? escapeHtml_(field.text) : field.text;
    return '<div style="position:absolute;left:' + left + 'px;top:' + top +
      'px;font-size:' + fontSize + 'px;font-family:\'Noto Sans JP\', \"Noto Sans\", sans-serif;font-weight:' + weight +
      ';text-align:' + align + ';white-space:nowrap;">' + content + '</div>';
  }).join('');

  const html = '<!doctype html><html><head><style>@page { size: ' + width + 'px ' + height + 'px; margin: 0; } body { margin: 0; padding: 0; }' +
    '</style></head><body>' +
    '<div style="position:relative;width:' + width + 'px;height:' + height + 'px;background:url(' + backgroundUrl +
    ') center center / contain no-repeat;">' + textLayer + '</div></body></html>';

  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF);
  return blob;
}

function selectBillingTemplateBlob_(templates, useCombined) {
  const id = useCombined ? templates.combined : templates.standard;
  if (!id) {
    throw new Error('請求書テンプレートIDが設定されていません');
  }
  const file = DriveApp.getFileById(id);
  return file.getBlob();
}

function buildCombinedBillingPdfBlob_(item, billingMonth, templates) {
  const carryOver = normalizeBillingCarryOver_(item);
  const total = normalizeBillingAmount_(item);
  const templateBlob = selectBillingTemplateBlob_(templates, carryOver > 0 || (item && item.shouldCombine));
  const pngBlob = templateBlob.getAs(MimeType.PNG);
  const overlay = {
    nameKanji: item && item.nameKanji ? item.nameKanji : '',
    address: resolveBillingAddress_(item),
    billingMonth: billingMonth ? String(billingMonth) : '',
    combinedTotal: formatBillingCurrency_(total),
    carryOver: carryOver ? formatBillingCurrency_(carryOver) : '',
    combineNote: '未入金があるため合算請求となります'
  };
  return createBillingPdfOverlay_(pngBlob, overlay);
}

function generateCombinedBillingPdfs(billingJson, options) {
  const opts = options || {};
  const billingMonth = opts.billingMonth || (Array.isArray(billingJson) && billingJson.length && billingJson[0].billingMonth) || '';
  const targets = (billingJson || []).filter(item => item && item.shouldCombine);
  const templates = resolveBillingPdfTemplates_(opts);
  const results = [];
  let folder = null;
  try {
    folder = getParentFolder_();
  } catch (e) {
    folder = null;
  }

  targets.forEach(item => {
    const title = '合算請求書_' + billingMonth + '_' + (item.patientId || '');
    const pdfBlob = buildCombinedBillingPdfBlob_(item, billingMonth, templates).setName(title + '.pdf');
    const pdfFile = folder ? folder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);
    results.push({
      patientId: item.patientId,
      fileId: pdfFile.getId(),
      url: pdfFile.getUrl(),
      name: pdfFile.getName()
    });
  });

  return { billingMonth, count: results.length, files: results };
}

function generateBillingOutputs(billingJson, options) {
  const excel = createBillingExcelFile(billingJson, options);
  const csv = createBillingCsvFile(billingJson, options);
  const history = appendBillingHistoryRows(billingJson, { billingMonth: excel.billingMonth || csv.billingMonth, memo: options && options.note });
  return { excel, csv, history };
}
