/***** Output layer: billing Excel/CSV/history generation *****/

const BILLING_OUTPUT_HEADER = [
  '請求月',
  '施術録番号',
  '氏名',
  'フリガナ',
  '保険区分',
  '負担割合',
  '施術回数',
  '単価',
  '合計金額',
  '請求金額',
  '銀行コード',
  '支店コード',
  '口座番号',
  '入金ステータス',
  '未入金額',
  '請求総額',
  '新規'
];

function formatBillingOutputRows(billingJson) {
  if (!Array.isArray(billingJson)) {
    throw new Error('請求データが不正です');
  }
  const sorted = billingJson.slice().sort((a, b) => {
    const aid = Number(a && a.patientId);
    const bid = Number(b && b.patientId);
    if (isNaN(aid) && isNaN(bid)) return 0;
    if (isNaN(aid)) return 1;
    if (isNaN(bid)) return -1;
    return aid - bid;
  });

  const rows = sorted.map(item => [
    item && item.billingMonth ? item.billingMonth : '',
    item && item.patientId ? item.patientId : '',
    item && item.nameKanji ? item.nameKanji : '',
    item && item.nameKana ? item.nameKana : '',
    item && item.insuranceType ? item.insuranceType : '',
    item && item.burdenRate != null ? item.burdenRate : '',
    item && item.visitCount != null ? item.visitCount : '',
    item && item.unitPrice != null ? item.unitPrice : '',
    item && item.total != null ? item.total : '',
    item && item.billingAmount != null ? item.billingAmount : '',
    item && item.bankCode ? item.bankCode : '',
    item && item.branchCode ? item.branchCode : '',
    item && item.accountNumber ? item.accountNumber : '',
    item && item.bankStatus ? item.bankStatus : '',
    item && item.carryOverAmount != null ? item.carryOverAmount : '',
    item && item.grandTotal != null ? item.grandTotal : '',
    item && item.isNew ? 1 : 0
  ]);

  return [BILLING_OUTPUT_HEADER].concat(rows);
}

function createBillingOutputSheet_(rows, sheetName) {
  const workbook = ss();
  const targetName = sheetName || '請求データ出力';
  let sheet = workbook.getSheetByName(targetName);
  if (!sheet) {
    sheet = workbook.insertSheet(targetName);
  } else {
    sheet.clear();
  }
  if (rows.length) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return sheet;
}

function createBillingExcelFile(billingJson, options) {
  const opts = options || {};
  const billingMonth = opts.billingMonth || (Array.isArray(billingJson) && billingJson.length && billingJson[0].billingMonth) || '';
  const baseName = opts.fileName || (billingMonth ? '請求データ_' + billingMonth : '請求データ');
  const rows = formatBillingOutputRows(billingJson);
  const temp = SpreadsheetApp.create(baseName);
  const sheet = temp.getSheets()[0];
  sheet.setName('請求データ');
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
  const baseName = opts.fileName || (billingMonth ? '請求データ_' + billingMonth : '請求データ');
  const rows = formatBillingOutputRows(billingJson);
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
    rowCount: rows.length ? rows.length - 1 : 0,
    billingMonth
  };
}

function ensureBillingHistorySheet_() {
  const SHEET_NAME = '請求出力履歴';
  const workbook = ss();
  let sheet = workbook.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 8).setValues([[
      '出力日時',
      '請求月',
      'レコード数',
      'Excel ファイルID',
      'Excel URL',
      'CSV ファイルID',
      'CSV URL',
      'メモ'
    ]]);
  }
  return sheet;
}

function recordBillingOutputHistory(params) {
  const sheet = ensureBillingHistorySheet_();
  const row = [
    new Date(),
    params && params.billingMonth ? params.billingMonth : '',
    params && params.rowCount ? params.rowCount : 0,
    params && params.excelFileId ? params.excelFileId : '',
    params && params.excelUrl ? params.excelUrl : '',
    params && params.csvFileId ? params.csvFileId : '',
    params && params.csvUrl ? params.csvUrl : '',
    params && params.note ? params.note : ''
  ];
  sheet.insertRows(2, 1);
  sheet.getRange(2, 1, 1, row.length).setValues([row]);
  return {
    billingMonth: row[1],
    rowNumber: 2
  };
}

function generateBillingOutputs(billingJson, options) {
  const excel = createBillingExcelFile(billingJson, options);
  const csv = createBillingCsvFile(billingJson, options);
  const history = recordBillingOutputHistory({
    billingMonth: excel.billingMonth || csv.billingMonth,
    rowCount: excel.rowCount,
    excelFileId: excel.fileId,
    excelUrl: excel.url,
    csvFileId: csv.fileId,
    csvUrl: csv.url,
    note: options && options.note ? options.note : ''
  });
  return { excel, csv, history };
}
