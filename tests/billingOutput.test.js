const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const billingOutputCode = fs.readFileSync(path.join(__dirname, '../src/output/billingOutput.js'), 'utf8');

function createContext() {
  return {
    console,
    MimeType: {
      GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet',
      MICROSOFT_EXCEL: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      PDF: 'application/pdf'
    }
  };
}

function createFakeFile(mimeType, tracker) {
  const blobTracker = tracker || { getAsCalled: false, setNameCalledWith: null };
  const blob = {
    getAs: () => {
      blobTracker.getAsCalled = true;
      return {
        setName: name => {
          blobTracker.setNameCalledWith = name;
          return { name };
        }
      };
    },
    setName: name => {
      blobTracker.setNameCalledWith = name;
      return { name };
    }
  };

  return {
    getMimeType: () => mimeType,
    getBlob: () => blob,
    tracker: blobTracker
  };
}

function testRejectsPdfBlobConversion() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const file = createFakeFile(context.MimeType.PDF);

  assert.throws(
    () => context.convertSpreadsheetToExcelBlob_(file, 'test'),
    /スプレッドシート以外のファイルをExcelに変換することはできません/,
    'PDF Blob が Excel 変換に渡された場合は例外を投げる'
  );
  assert.strictEqual(file.tracker.getAsCalled, false, 'PDF Blob では getAs が呼び出されない');
}

function testSpreadsheetBlobIsConverted() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const tracker = { getAsCalled: false, setNameCalledWith: null };
  const file = createFakeFile(context.MimeType.GOOGLE_SHEETS, tracker);

  const result = context.convertSpreadsheetToExcelBlob_(file, 'export_name');
  assert.deepStrictEqual(result, { name: 'export_name.xlsx' }, 'Excel 変換結果が返却される');
  assert.strictEqual(tracker.getAsCalled, true, 'Spreadsheet Blob では getAs が呼び出される');
  assert.strictEqual(tracker.setNameCalledWith, 'export_name.xlsx', 'setName が適切なファイル名で呼ばれる');
}

function testExcelBlobIsReturnedWithoutConversion() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const tracker = { getAsCalled: false, setNameCalledWith: null };
  const file = createFakeFile(context.MimeType.MICROSOFT_EXCEL, tracker);

  const result = context.convertSpreadsheetToExcelBlob_(file, 'already_excel');
  assert.deepStrictEqual(result, { name: 'already_excel.xlsx' }, 'Excel Blob はそのまま返却される');
  assert.strictEqual(tracker.getAsCalled, false, 'Excel Blob では getAs が呼び出されない');
  assert.strictEqual(tracker.setNameCalledWith, 'already_excel.xlsx', '既存の Excel にも名称設定が行われる');
}

function testCustomUnitPriceForSelfPaidInvoice() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const { calculateInvoiceChargeBreakdown_ } = context;
  assert.strictEqual(typeof calculateInvoiceChargeBreakdown_, 'function', '請求額計算の関数が定義されている');

  const result = calculateInvoiceChargeBreakdown_({
    insuranceType: '自費',
    unitPrice: 5000,
    burdenRate: '',
    visitCount: 2,
    carryOverAmount: 1000
  });

  assert.strictEqual(result.treatmentUnitPrice, 5000, '自費の場合はカスタム単価を使用する');
  assert.strictEqual(result.treatmentAmount, 10000, '単価と訪問回数から施術料が算出される');
  assert.strictEqual(result.transportAmount, 66, '自費でも交通費が計上される');
  assert.strictEqual(result.grandTotal, 11066, '繰越分も含めて合計が算出される');
}

function run() {
  testRejectsPdfBlobConversion();
  testSpreadsheetBlobIsConverted();
  testExcelBlobIsReturnedWithoutConversion();
  testCustomUnitPriceForSelfPaidInvoice();
  console.log('billingOutput blob guard tests passed');
}

run();
