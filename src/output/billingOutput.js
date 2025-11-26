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

function createBillingExcelFile(billingJson, options) {
  const opts = options || {};
  const billingMonth = opts.billingMonth || (Array.isArray(billingJson) && billingJson.length && billingJson[0].billingMonth) || '';
  const baseName = opts.fileName || (billingMonth ? '請求一覧_' + billingMonth : '請求一覧');
  const rows = buildBillingExcelRows_(billingJson);
  const temp = SpreadsheetApp.create(baseName);
  const sheet = temp.getSheets()[0];
  sheet.setName('請求一覧');
  if (rows.length) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }

  const tempFile = DriveApp.getFileById(temp.getId());
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
    const status = statusMap[pid] && statusMap[pid].bankStatus ? statusMap[pid].bankStatus : null;
    if (!status) return;
    const newRow = row.slice();
    newRow[8] = status;
    newRow[9] = new Date();
    updates.push({ rowNumber: idx + 2, values: newRow });
  });
  updates.forEach(update => {
    sheet.getRange(update.rowNumber, 1, 1, update.values.length).setValues([update.values]);
  });
  return { billingMonth, updated: updates.length };
}

function generateCombinedBillingPdfs(billingJson, options) {
  const opts = options || {};
  const billingMonth = opts.billingMonth || (Array.isArray(billingJson) && billingJson.length && billingJson[0].billingMonth) || '';
  const targets = (billingJson || []).filter(item => item && item.shouldCombine);
  const results = [];
  let folder = null;
  try {
    folder = getParentFolder_();
  } catch (e) {
    folder = null;
  }

  targets.forEach(item => {
    const title = '合算請求書_' + billingMonth + '_' + (item.patientId || '');
    const doc = DocumentApp.create(title);
    const body = doc.getBody();
    body.appendParagraph('合算請求書');
    body.appendParagraph('請求月: ' + billingMonth);
    body.appendParagraph('患者ID: ' + (item.patientId || ''));
    body.appendParagraph('氏名: ' + (item.nameKanji || ''));
    body.appendParagraph('請求金額: ' + normalizeBillingAmount_(item));
    body.appendParagraph('未入金額: ' + (item && item.carryOverAmount != null ? item.carryOverAmount : 0));
    body.appendParagraph('未入金分を含む合算請求となります。');

    const docFile = DriveApp.getFileById(doc.getId());
    if (folder) {
      folder.addFile(docFile);
      DriveApp.getRootFolder().removeFile(docFile);
    }
    const pdfBlob = docFile.getBlob().getAs(MimeType.PDF).setName(title + '.pdf');
    const pdfFile = folder ? folder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);
    docFile.setTrashed(true);
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
