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

function testBillingAmountFallsBackToTotals() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const { normalizeBillingAmount_ } = context;

  const amountFromParts = normalizeBillingAmount_({
    billingAmount: '2,000',
    transportAmount: '330',
    carryOverAmount: 500
  });
  assert.strictEqual(amountFromParts, 2830, '請求額・交通費・繰越の合算を返す');

  const amountFromTotal = normalizeBillingAmount_({
    total: 2500,
    carryOverAmount: 400,
    carryOverFromHistory: 100
  });
  assert.strictEqual(amountFromTotal, 3000, 'total があれば繰越を加算して返す');
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

  assert.strictEqual(result.treatmentUnitPrice, 5000, '自費でも手動入力の単価が優先される');
  assert.strictEqual(result.treatmentAmount, 10000, '自費で単価が指定された場合は施術料を計上する');
  assert.strictEqual(result.transportAmount, 66, '自費で単価を入力した場合でも交通費を計上する');
  assert.strictEqual(result.grandTotal, 11066, '手動単価と交通費・繰越を合算して出力する');
}

function testFullWidthInputsAreNormalized() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const breakdown = context.calculateInvoiceChargeBreakdown_({
    insuranceType: '自費',
    unitPrice: '５,０００',
    burdenRate: '',
    visitCount: '２',
    carryOverAmount: '１，０００'
  });

  assert.strictEqual(breakdown.visits, 2, '全角の回数も計上される');
  assert.strictEqual(breakdown.treatmentUnitPrice, 5000, '全角入力の単価も自費で優先される');
  assert.strictEqual(breakdown.transportAmount, 66, '全角入力でも交通費が計上される');
  assert.strictEqual(breakdown.grandTotal, 11066, '全角入力でも合計に反映される');
}

function testSelfPaidInvoiceStaysZeroWithoutManualUnitPrice() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const breakdown = context.calculateInvoiceChargeBreakdown_({
    insuranceType: '自費',
    burdenRate: '',
    visitCount: 3,
    carryOverAmount: 500
  });

  assert.strictEqual(breakdown.treatmentUnitPrice, 0, '単価未設定の自費は施術料0円となる');
  assert.strictEqual(breakdown.transportAmount, 0, '単価未設定なら交通費は計上されない');
  assert.strictEqual(breakdown.grandTotal, 500, '繰越のみが合計に残る');
}

function testSelfPaidInvoiceDoesNotRoundManualUnitPrice() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const breakdown = context.calculateInvoiceChargeBreakdown_({
    insuranceType: '自費',
    unitPrice: 3333,
    visitCount: 1,
    carryOverAmount: 0
  });

  assert.strictEqual(breakdown.treatmentAmount, 3333, '自費の手動単価は四捨五入せずに計上する');
  assert.strictEqual(breakdown.grandTotal, 3366, '施術料と交通費の合計をそのまま出力する');
}

function testInsuranceBillingIsRoundedToNearestTen() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const breakdown = context.calculateInvoiceChargeBreakdown_({
    insuranceType: '鍼灸',
    burdenRate: 1,
    visitCount: 7,
    carryOverAmount: 0
  });

  assert.strictEqual(breakdown.treatmentAmount, 2920, '施術料は10円単位で四捨五入される');
  assert.strictEqual(breakdown.transportAmount, 231, '交通費は回数分計上される');
  assert.strictEqual(breakdown.grandTotal, 3151, '合計も四捨五入後の施術料を利用する');
}

function testWelfareBillingStillAddsTransport() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const breakdown = context.calculateInvoiceChargeBreakdown_({
    insuranceType: '生保',
    visitCount: 5,
    carryOverAmount: 0
  });

  assert.strictEqual(breakdown.treatmentAmount, 0, '生保は施術料が0円のまま');
  assert.strictEqual(breakdown.transportAmount, 0, '生保では交通費を請求しない');
  assert.strictEqual(breakdown.grandTotal, 0, '交通費なしの場合は合計も0円となる');
}

function testMassageBillingDoesNotChargeTransport() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const breakdown = context.calculateInvoiceChargeBreakdown_({
    insuranceType: 'マッサージ',
    visitCount: 4,
    carryOverAmount: 200
  });

  assert.strictEqual(breakdown.treatmentAmount, 0, 'マッサージは施術料が0円のまま');
  assert.strictEqual(breakdown.transportAmount, 0, 'マッサージでは交通費を請求しない');
  assert.strictEqual(breakdown.grandTotal, 200, '繰越のみの場合は交通費なしで合計される');
}

function testCarryOverHistoryIsIncluded() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const breakdown = context.calculateInvoiceChargeBreakdown_({
    insuranceType: '鍼灸',
    burdenRate: 1,
    visitCount: 1,
    carryOverAmount: 500,
    carryOverFromHistory: 200
  });

  assert.strictEqual(breakdown.treatmentAmount, 420, '施術料は四捨五入後の負担額で計算される');
  assert.strictEqual(breakdown.grandTotal, 1153, '未回収分も繰越に合算される');
}

function run() {
  testRejectsPdfBlobConversion();
  testSpreadsheetBlobIsConverted();
  testExcelBlobIsReturnedWithoutConversion();
  testBillingAmountFallsBackToTotals();
  testCustomUnitPriceForSelfPaidInvoice();
  testFullWidthInputsAreNormalized();
  testSelfPaidInvoiceStaysZeroWithoutManualUnitPrice();
  testSelfPaidInvoiceDoesNotRoundManualUnitPrice();
  testInsuranceBillingIsRoundedToNearestTen();
  testWelfareBillingStillAddsTransport();
  testMassageBillingDoesNotChargeTransport();
  testCarryOverHistoryIsIncluded();
  console.log('billingOutput blob guard tests passed');
}

run();
