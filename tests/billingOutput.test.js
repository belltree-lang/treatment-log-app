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

function createExportContext(overrides = {}) {
  const base = {
    console,
    billingLogger_: { log: () => {} },
    normalizePreparedBilling_: payload => {
      if (!payload) return null;
      const billingJson = Array.isArray(payload.billingJson) ? payload.billingJson : [];
      return Object.assign({ billingJson }, payload);
    },
    logPreparedBankPayloadStatus_: () => {},
    normalizeZeroOneFlag_: value => (value === 1 || value === '1' || value === true ? 1 : 0),
    getBillingSourceData: () => ({
      billingMonth: '202501',
      bankInfoByName: {},
      patients: {},
      bankStatuses: {}
    }),
    generateBillingJsonFromSource: () => ([{ billingMonth: '202501', patientId: 'P001', nameKanji: 'テスト' }])
  };

  const context = Object.assign({}, base, overrides);
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  if (overrides.exportBankTransferRows_) {
    context.exportBankTransferRows_ = overrides.exportBankTransferRows_;
  }

  return context;
}

class FakeSheet {
  constructor(headers, rows = []) {
    this.values = [headers.slice(), ...rows];
  }

  getLastRow() {
    return this.values.length;
  }

  getLastColumn() {
    return this.values[0] ? this.values[0].length : 0;
  }

  getRange(row, col, numRows = 1, numCols = 1) {
    const startRow = row - 1;
    const startCol = col - 1;
    const endRow = startRow + numRows;
    const endCol = startCol + numCols;

    const ensureSize = () => {
      while (this.values.length < endRow) {
        this.values.push(new Array(this.getLastColumn() || endCol).fill(''));
      }
      this.values.forEach(r => {
        while (r.length < endCol) r.push('');
      });
    };

    const slice = () => this.values
      .slice(startRow, endRow)
      .map(r => r.slice(startCol, endCol));

    return {
      getDisplayValues: () => slice(),
      getValues: () => slice(),
      setValues: (vals) => {
        ensureSize();
        vals.forEach((r, i) => {
          this.values[startRow + i].splice(startCol, numCols, ...r);
        });
      },
      setValue: (value) => {
        ensureSize();
        this.values[startRow][startCol] = value;
      },
      clearContent: () => {
        ensureSize();
        for (let r = startRow; r < endRow; r += 1) {
          for (let c = startCol; c < endCol; c += 1) {
            this.values[r][c] = '';
          }
        }
      }
    };
  }
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

function testReceiptVisibilityRespectsBankFlagsAndStatus() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const { resolveInvoiceReceiptDisplay_ } = context;

  const defaultStatus = resolveInvoiceReceiptDisplay_({ billingMonth: '202501', hasPreviousPrepared: true });
  assert.strictEqual(defaultStatus.showReceipt, true, '銀行引落シートがあれば前月領収書を表示する');
  assert.deepStrictEqual(Array.from(defaultStatus.receiptMonths || []), ['202412'], '前月の領収書を作成する');

  const withoutPreviousSheet = resolveInvoiceReceiptDisplay_({ billingMonth: '202501', hasPreviousPrepared: false });
  assert.strictEqual(withoutPreviousSheet.showReceipt, false, '前月請求が無ければ領収書を非表示にする');

  const withUnpaidStatus = resolveInvoiceReceiptDisplay_({
    billingMonth: '202501',
    hasPreviousPrepared: true,
    receiptStatus: 'UNPAID'
  });
  assert.strictEqual(withUnpaidStatus.showReceipt, false, '領収ステータスが UNPAID のときは非表示にする');

  const withSkipReceipt = resolveInvoiceReceiptDisplay_({
    billingMonth: '202501',
    hasPreviousPrepared: true,
    skipReceipt: true
  });
  assert.strictEqual(withSkipReceipt.showReceipt, false, 'skipReceipt が指定された場合は非表示にする');
}

function testInvoiceTemplateSwitchesAggregateModeForUnpaid() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const { buildInvoiceTemplateData_ } = context;

  const aggregate = buildInvoiceTemplateData_({
    billingMonth: '202502',
    aggregateStatus: 'confirmed',
    aggregateTargetMonths: ['202412', '202501', '202502'],
    receiptMonths: ['202412', '202501', '202502'],
    skipReceipt: true,
    hasPreviousReceiptSheet: true
  });

  assert.strictEqual(aggregate.isAggregateInvoice, true, '合算対象があれば合算モードになる');
  assert.strictEqual(aggregate.invoiceMode, 'aggregate', 'モード名を保持する');
  assert.strictEqual(aggregate.chargeMonthLabel, '2025年02月', '請求月のみを表示する');
  assert.strictEqual(aggregate.showReceipt, false, '合算請求時は領収書を出力しない');

  const standard = buildInvoiceTemplateData_({ billingMonth: '202502', hasPreviousPrepared: true });
  assert.strictEqual(standard.isAggregateInvoice, false, '未回収が無ければ通常モード');
  assert.strictEqual(standard.invoiceMode, 'standard', '通常モードを示す');
  assert.strictEqual(standard.chargeMonthLabel, '2025年02月', '請求月のみを表示する');
  assert.strictEqual(standard.showReceipt, true, '通常モードでは領収書を表示する');
}

function testInvoiceTemplateIgnoresFallbackReceiptMonthForAggregate() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const { buildInvoiceTemplateData_ } = context;

  const standard = buildInvoiceTemplateData_({ billingMonth: '202511', hasPreviousPrepared: true });
  assert.strictEqual(standard.isAggregateInvoice, false, 'フォールバック領収月のみでは合算モードにしない');
  assert.strictEqual(standard.invoiceMode, 'standard', 'フォールバックでは通常モードを維持する');
  assert.strictEqual(standard.aggregateMonthTotals.length, 0, '合算内訳は生成しない');
  assert.strictEqual(standard.chargeMonthLabel, '2025年11月', '請求月表示は billingMonth を優先する');
  const trace = standard.aggregateDecisionTrace;
  assert.ok(trace, '合算判定トレースを含める');
  assert.strictEqual(trace.receiptMonthsSource, 'fallback', '領収月のソースを明示する');
  assert.deepStrictEqual(Array.from(trace.decisionSources || []), [], 'フォールバック領収月は判定ソースに含めない');
}

function testReceiptDisplayFallsBackToPreviousMonthWhenDefault() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const display = context.resolveInvoiceReceiptDisplay_({
    billingMonth: '202411'
  });

  assert.strictEqual(display.visible, true, '請求月のみ指定でも領収書を表示する');
  assert.deepStrictEqual(Array.from(display.receiptMonths || []), ['202410'], '前月を単月領収書として表示する');
  assert.strictEqual(display.receiptMonthsSource, 'fallback', '表示ソースをフォールバックとして保持する');
  assert.deepStrictEqual(Array.from(display.explicitReceiptMonths || []), [], '明示指定は無いままにする');
}

function testAggregateDecisionIgnoresPreviousReceiptAmount() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const aggregate = context.buildInvoiceTemplateData_({
    billingMonth: '202501',
    previousReceiptAmount: 5000,
    receiptMonths: ['202412']
  });

  assert.strictEqual(aggregate.isAggregateInvoice, false, '金額のみでは合算扱いにしない');
}

function testAggregateInvoiceHidesReceiptWhenSkipped() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const display = context.resolveInvoiceReceiptDisplay_({
    billingMonth: '202501',
    hasPreviousPrepared: true,
    receiptMonths: ['202412', '202501'],
    skipReceipt: true
  });

  assert.strictEqual(display.visible, false, '合算請求時は領収書を表示しない');
}

function testAggregateStatusDoesNotFinalizeWithoutConfirmation() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const { buildInvoiceTemplateData_ } = context;

  const aggregate = buildInvoiceTemplateData_({
    billingMonth: '202501',
    aggregateTargetMonths: ['202411', '202412'],
    aggregateStatus: 'scheduled'
  });
  assert.strictEqual(aggregate.isAggregateInvoice, true, '明示的な合算対象があれば合算判定は true');
  assert.strictEqual(aggregate.aggregateConfirmed, false, 'scheduled では確定扱いにしない');
  assert.strictEqual(aggregate.finalized, false, '合算でも confirmed でなければ確定扱いにしない');

  const confirmed = buildInvoiceTemplateData_({
    billingMonth: '202501',
    aggregateTargetMonths: ['202411', '202412'],
    aggregateStatus: 'confirmed'
  });
  assert.strictEqual(confirmed.aggregateConfirmed, true, 'confirmed のときのみ確定扱いにする');
  assert.strictEqual(confirmed.finalized, true, 'confirmed なら確定扱い');
}

function testPreviousReceiptSettlementRequiresExplicitStatus() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const { isPreviousReceiptSettled_, buildInvoicePreviousReceipt_, buildInvoiceTemplateData_ } = context;
  assert.strictEqual(isPreviousReceiptSettled_({}), false, 'ステータス不明なら未確定扱い');
  assert.strictEqual(isPreviousReceiptSettled_({ previousReceiptStatus: 'SETTLED' }), true, 'previousReceiptStatus=SETTLED を settled とみなす');
  assert.strictEqual(isPreviousReceiptSettled_({ receiptStatus: 'settled' }), true, '領収ステータスでも settled を認める');
  assert.strictEqual(isPreviousReceiptSettled_({ previousReceiptStatus: 'pending' }), false, 'その他は未確定扱い');

  const receipt = buildInvoicePreviousReceipt_({ billingMonth: '202501', previousReceiptStatus: 'SETTLED' });
  assert.strictEqual(receipt.settled, true, '前月領収情報に settled フラグを含める');

  const finalizedByReceipt = buildInvoiceTemplateData_({
    billingMonth: '202501',
    previousReceiptStatus: 'SETTLED'
  });
  assert.strictEqual(finalizedByReceipt.finalized, true, '前月領収ステータスが SETTLED のとき確定扱いにする');

  const injectedSettled = buildInvoiceTemplateData_({
    billingMonth: '202501',
    previousReceipt: { settled: true }
  });
  assert.strictEqual(injectedSettled.finalized, false, '明示的なステータスなしの settled=true では確定扱いにしない');
}

function testPreviousReceiptVisibilityFollowsReceiptDecision() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const { buildInvoiceTemplateData_ } = context;

  const payable = buildInvoiceTemplateData_({ billingMonth: '202501', receiptStatus: 'PAID', hasPreviousPrepared: true });
  assert.strictEqual(payable.previousReceipt.visible, true, '銀行引落シートがあれば前月領収書も表示する');

  const onHold = buildInvoiceTemplateData_({ billingMonth: '202501', receiptStatus: 'HOLD', hasPreviousPrepared: true });
  assert.strictEqual(onHold.previousReceipt.visible, false, 'HOLD の場合は領収書を非表示にする');
}

function testPreviousReceiptIsHiddenWhenPreviousPreparedMissing() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const { buildInvoiceTemplateData_ } = context;

  const data = buildInvoiceTemplateData_({ billingMonth: '202501', receiptStatus: 'PAID', hasPreviousPrepared: false });

  assert.strictEqual(data.previousReceipt.visible, false, '前月請求が未生成なら前月領収書を非表示にする');
}

function testAggregateTemplateUsesExplicitMonths() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const { buildAggregateInvoiceTemplateData_ } = context;

  const data = buildAggregateInvoiceTemplateData_(
    { billingMonth: '202511', patientId: '001' },
    ['202509', '202510']
  );

  assert.deepStrictEqual(
    Array.from(data.receiptMonths || []),
    ['202509', '202510'],
    '指定した合算月のみを請求対象に使う'
  );
  assert.strictEqual(data.chargeMonthLabel, '2025年11月', '請求月の表示は billingMonth を基準にする');
  assert.strictEqual(data.monthLabel, '2025年11月（合算）', '合算時もベースは請求月ラベル');
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

function testReceiptStatusIsOverwrittenInHistory() {
  const context = createExportContext({
    ss: () => ({
      getSheetByName: name => (name === '請求履歴' ? sheet : null),
      insertSheet: () => sheet
    })
  });

  const existingRow = [
    '202501',
    '001',
    '山田太郎',
    1000,
    0,
    1000,
    500,
    500,
    'OK',
    new Date('2025-02-01'),
    '初期メモ',
    'UNPAID',
    '202503'
  ];

  const headers = context.BILLING_HISTORY_HEADERS || [
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
    'memo',
    'receiptStatus',
    'aggregateUntilMonth'
  ];

  const sheet = new FakeSheet(headers, [existingRow]);

  context.appendBillingHistoryRows([
    {
      billingMonth: '202501',
      patientId: '001',
      nameKanji: '山田太郎',
      receiptStatus: 'AGGREGATE',
      aggregateUntilMonth: ''
    }
  ], { billingMonth: '202501' });

  const { columns } = context.resolveBillingHistoryColumns_(sheet);
  const row = sheet.values[1];

  assert.strictEqual(row[columns.receiptStatus - 1], 'AGGREGATE', '既存の領収ステータスを上書きする');
  assert.strictEqual(row[columns.aggregateUntilMonth - 1], '', '集計終了月のリセットも反映される');
}

function testInsuranceBillingUsesYenRounding() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const breakdown = context.calculateInvoiceChargeBreakdown_({
    insuranceType: '鍼灸',
    burdenRate: 1,
    visitCount: 7,
    carryOverAmount: 0
  });

  assert.strictEqual(breakdown.treatmentAmount, 2919, '施術料は円単位で計算される');
  assert.strictEqual(breakdown.transportAmount, 231, '交通費は回数分計上される');
  assert.strictEqual(breakdown.grandTotal, 3150, '合計も円単位で計算された施術料を利用する');
}

function testTwoTenthBurdenKeepsYenPrecision() {
  const context = createContext();
  vm.createContext(context);
  vm.runInContext(billingOutputCode, context);

  const breakdown = context.calculateInvoiceChargeBreakdown_({
    insuranceType: '鍼灸',
    burdenRate: 2,
    unitPrice: 4170,
    visitCount: 6,
    carryOverAmount: 0
  });

  assert.strictEqual(breakdown.treatmentAmount, 5004, '2割負担でも円単位で計算される');
  assert.strictEqual(breakdown.transportAmount, 198, '交通費は回数分計上される');
  assert.strictEqual(breakdown.grandTotal, 5202, '施術料と交通費の合計が正しく算出される');
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

  assert.strictEqual(breakdown.treatmentAmount, 417, '施術料は円単位の負担額で計算される');
  assert.strictEqual(breakdown.grandTotal, 1150, '未回収分も繰越に合算される');
}

function testBankExportRejectsNullPreparedPayload() {
  const context = createExportContext();
  assert.throws(() => {
    context.exportBankTransferDataForPrepared_(null);
  }, /請求データが未生成/);
}

function testBankExportUsesNormalizedBillingJsonWhenMissingFromPayload() {
  const normalizedBillingJson = [{ billingMonth: '202501', patientId: 'P001', nameKanji: '正規化済み' }];
  const bankInfoByName = { normalized: { bankCode: '0001', branchCode: '002', accountNumber: '1234567' } };
  const patients = { P001: { bankCode: '0001', branchCode: '002', accountNumber: '1234567' } };
  const bankFlagsByPatient = { P001: { ae: true, af: false } };
  const buildCalls = [];
  const exportCalls = [];

  const context = createExportContext({
    normalizePreparedBilling_: payload => Object.assign({
      billingMonth: payload.billingMonth,
      billingJson: normalizedBillingJson,
      bankInfoByName,
      patients,
      bankStatuses: {},
      bankFlagsByPatient
    }, payload),
    logPreparedBankPayloadStatus_: () => {}
  });

  context.buildBankTransferRowsForBilling_ = (billingJson, bankInfoByNameParam, patientMapParam, billingMonth, bankStatuses) => {
    buildCalls.push({ billingJson, bankInfoByName: bankInfoByNameParam, patientMap: patientMapParam, billingMonth, bankStatuses });
    return { billingMonth, rows: [{ patientId: 'P001' }], total: billingJson.length, passed: billingJson.length, skipped: 0, skipReasons: {} };
  };
  context.exportBankTransferRows_ = (billingMonth, rows) => {
    exportCalls.push({ billingMonth, rows });
    return { inserted: rows.length, skipped: 0 };
  };

  const result = context.exportBankTransferDataForPrepared_({ billingMonth: '202501' });

  assert.strictEqual(buildCalls.length, 1, '正規化済みの billingJson で銀行CSV構築に進む');
  assert.deepStrictEqual(buildCalls[0].billingJson, [{
    billingMonth: '202501',
    patientId: 'P001',
    nameKanji: '正規化済み',
    bankFlags: { ae: true, af: false }
  }], 'billingJson は bankFlags を付与した上で渡される');
  assert.strictEqual(exportCalls.length, 1, 'エクスポート処理まで到達する');
  assert.strictEqual(result.rows.length, 1, '正規化済みデータから行が生成される');
}

function testBankExportRejectsInvalidBillingJsonShape() {
  const context = createExportContext();
  assert.throws(() => {
    context.exportBankTransferDataForPrepared_({ billingMonth: '202501', billingJson: 'not-array' });
  }, /形式が不正/);
}

function testBankExportReturnsEmptyWhenNoRows() {
  const context = createExportContext({
    exportBankTransferRows_: () => assert.fail('export should not be called for empty payload')
  });
  const result = context.exportBankTransferDataForPrepared_({ billingMonth: '202501', billingJson: [] });

  assert.strictEqual(result.billingMonth, '202501');
  assert.ok(Array.isArray(result.rows) && result.rows.length === 0, 'rows should be an empty array');
  assert.strictEqual(result.inserted, 0);
  assert.strictEqual(result.skipped, 0);
  assert.match(result.message, /請求対象はありません/);
}

function testBankCodesAreNormalizedBeforeValidation() {
  const context = createExportContext();
  const { buildBankTransferRowsForBilling_, normalizeBillingNameKey_ } = context;

  const billingJson = [{ billingMonth: '202502', patientId: '001', nameKanji: '山田太郎' }];
  const patientMap = { '001': { bankCode: '12', branchCode: '3', accountNumber: '45678' } };
  const bankInfoByName = {
    [normalizeBillingNameKey_('山田太郎')]: {
      bankCode: '1-23',
      branchCode: '45',
      accountNumber: '6789'
    }
  };

  const result = buildBankTransferRowsForBilling_(billingJson, bankInfoByName, patientMap, '202502', {});

  assert.strictEqual(result.skipped, 0, '正規化後の銀行情報はスキップされない');
  assert.strictEqual(result.rows.length, 1, '銀行CSVの行が生成される');

  const row = result.rows[0];
  assert.strictEqual(row.bankCode, '0123', '銀行コードは数字のみ・左ゼロ埋めで4桁になる');
  assert.strictEqual(row.branchCode, '045', '支店コードは数字のみ・左ゼロ埋めで3桁になる');
  assert.strictEqual(row.accountNumber, '0006789', '口座番号は数字のみ・左ゼロ埋めで7桁になる');
}

function testUnifiedNameKeyAcrossBankPatientAndBilling() {
  const context = createExportContext();
  const { buildBankTransferRowsForBilling_, normalizeBillingFullNameKey_ } = context;

  const billingJson = [{
    billingMonth: '202503',
    patientId: '0005',
    nameKanji: '山田　太郎',
    nameKana: 'ヤマダ タロウ',
    billingAmount: 5000
  }];

  const patientMap = {
    '0005': {
      patientId: '0005',
      nameKanji: '山田太郎',
      nameKana: 'ヤマダタロウ'
    }
  };

  const bankInfoByName = {
    [normalizeBillingFullNameKey_('山田太郎', 'ﾔﾏﾀﾞﾀﾛｳ')]: {
      nameKanji: '山田太郎',
      nameKana: 'ﾔﾏﾀﾞﾀﾛｳ',
      bankCode: '12 3',
      branchCode: '45',
      accountNumber: '6789'
    }
  };

  const result = buildBankTransferRowsForBilling_(billingJson, bankInfoByName, patientMap, '202503', {});

  assert.strictEqual(result.total, 1, '銀行データ・患者マスタ・billingJson を統一キーで突合する');
  assert.strictEqual(result.skipped, 0, '統一キーで突合したデータはスキップされない');
  assert.strictEqual(result.rows.length, 1, '突合に成功したデータから行が生成される');

  const row = result.rows[0];
  assert.strictEqual(row.patientId, '0005', '患者ID は billingJson から引き継がれる');
  assert.strictEqual(row.nameKanji, '山田太郎', '氏名（漢字）は正規化された銀行情報が使われる');
  assert.strictEqual(row.nameKana, 'ヤマダタロウ', '氏名（カナ）は正規化後の値が使われる');
  assert.strictEqual(row.bankCode, '0123', '銀行コードは統一キーでマージした情報から取得される');
  assert.strictEqual(row.branchCode, '045', '支店コードは正規化される');
  assert.strictEqual(row.accountNumber, '0006789', '口座番号は正規化される');
}

function testNameKeyNormalizationStripsSeparators() {
  const context = createExportContext();
  const { buildBankTransferRowsForBilling_, normalizeBillingFullNameKey_ } = context;

  const billingJson = [{
    billingMonth: '202504',
    patientId: '010',
    nameKanji: '山田太郎',
    nameKana: 'ヤマダタロウ',
    billingAmount: 8000
  }];

  const patientMap = {
    '010': {
      patientId: '010',
      nameKanji: '山田太郎',
      nameKana: 'ヤマダタロウ',
      bankCode: '12',
      branchCode: '34',
      accountNumber: '567890'
    }
  };

  const bankInfoByName = {
    [normalizeBillingFullNameKey_('山田・太郎', 'ﾔﾏﾀﾞ･ﾀﾛｳ')]: {
      nameKanji: '山田・太郎',
      nameKana: 'ﾔﾏﾀﾞ･ﾀﾛｳ',
      bankCode: '12',
      branchCode: '34',
      accountNumber: '567890'
    }
  };

  const result = buildBankTransferRowsForBilling_(billingJson, bankInfoByName, patientMap, '202504', {});

  assert.strictEqual(result.rows.length, 1, '区切り文字の有無に関わらず突合できる');
  assert.strictEqual(result.rows[0].nameKanji, '山田太郎', '漢字の区切り文字は正規化で除去される');
  assert.strictEqual(result.rows[0].nameKana, 'ヤマダタロウ', 'カナの区切り文字も正規化される');
}

function testBankRowsSkipWhenNormalizedLengthsAreInvalid() {
  const context = createExportContext();
  const { buildBankTransferRowsForBilling_ } = context;

  const billingJson = [{ billingMonth: '202502', patientId: '002', nameKanji: '銀行エラー' }];
  const patientMap = {
    '002': { bankCode: '12345', branchCode: '678', accountNumber: '123456789' }
  };

  const result = buildBankTransferRowsForBilling_(billingJson, {}, patientMap, '202502', {});

  assert.strictEqual(result.rows.length, 0, '不正な桁数のデータは行に含まれない');
  assert.strictEqual(result.skipped, 1, '不正データはスキップ数としてカウントされる');
  assert.strictEqual(result.total, 1, '総件数がカウントされる');
  assert.strictEqual(result.passed, 0, '有効データが0件であることが分かる');
  const skipSummary = Object.assign({}, result.skipReasons);
  assert.deepStrictEqual(skipSummary, {
    invalidBankCode: 1,
    invalidBranchCode: 0,
    invalidAccountNumber: 1
  }, 'スキップ理由の内訳が返却される');
}

function testNameKanaFallsBackToKanjiWhenEmpty() {
  const context = createExportContext();
  const { buildBankTransferRowsForBilling_ } = context;

  const billingJson = [{ billingMonth: '202502', patientId: '003', nameKanji: '山田太郎' }];
  const patientMap = {
    '003': { bankCode: '1', branchCode: '2', accountNumber: '3', nameKanji: '山田太郎' }
  };

  const result = buildBankTransferRowsForBilling_(billingJson, {}, patientMap, '202502', {});

  assert.strictEqual(result.rows.length, 1, '名義カナが空でも行が生成される');
  assert.strictEqual(result.rows[0].nameKana, '山田太郎', '名義カナが空の場合は漢字から代替生成される');
}

function testNameKanaIsNormalizedToFullWidth() {
  const context = createExportContext();
  const { buildBankTransferRowsForBilling_ } = context;

  const billingJson = [{ billingMonth: '202502', patientId: '004', nameKanji: '佐藤花子', nameKana: ' ﾊﾅｺ ' }];
  const patientMap = {
    '004': { bankCode: '1', branchCode: '2', accountNumber: '3', nameKanji: '佐藤花子' }
  };

  const result = buildBankTransferRowsForBilling_(billingJson, {}, patientMap, '202502', {});

  assert.strictEqual(result.rows.length, 1, '半角名義カナでも行が生成される');
  assert.strictEqual(result.rows[0].nameKana, 'ハナコ', '名義カナはNFKC変換とtrimで正規化される');
}

function testSkipMessageIncludesReasonsWhenAllRowsInvalid() {
  const context = createExportContext({
    exportBankTransferRows_: (billingMonth, rows) => ({ billingMonth, inserted: rows.length })
  });

  const prepared = {
    billingMonth: '202503',
    billingJson: [{ billingMonth: '202503', patientId: '005', nameKanji: '銀行不備' }],
    bankInfoByName: { placeholder: {} },
    patients: {
      '005': { bankCode: '12345', branchCode: '1234', accountNumber: '123456789' }
    },
    bankStatuses: {}
  };

  const result = context.exportBankTransferDataForPrepared_(prepared);

  assert.strictEqual(result.rows.length, 0, '全件スキップ時は行が空のままになる');
  assert.ok(result.message.includes('銀行CSVが生成されませんでした'), 'スキップ理由がメッセージに含まれる');
  assert.match(result.message, /総件数: 1/, '総件数がメッセージに含まれる');
  assert.match(result.message, /銀行コード不正: 1/, '銀行コードの不正件数が含まれる');
  assert.match(result.message, /支店コード不正: 1/, '支店コードの不正件数が含まれる');
  assert.match(result.message, /口座番号不正: 1/, '口座番号の不正件数が含まれる');
}

function run() {
  testRejectsPdfBlobConversion();
  testSpreadsheetBlobIsConverted();
  testExcelBlobIsReturnedWithoutConversion();
  testBillingAmountFallsBackToTotals();
  testCustomUnitPriceForSelfPaidInvoice();
  testFullWidthInputsAreNormalized();
  testSelfPaidInvoiceStaysZeroWithoutManualUnitPrice();
  testReceiptVisibilityRespectsBankFlagsAndStatus();
  testInvoiceTemplateSwitchesAggregateModeForUnpaid();
  testInvoiceTemplateIgnoresFallbackReceiptMonthForAggregate();
  testReceiptDisplayFallsBackToPreviousMonthWhenDefault();
  testAggregateDecisionIgnoresPreviousReceiptAmount();
  testAggregateInvoiceHidesReceiptWhenSkipped();
  testAggregateStatusDoesNotFinalizeWithoutConfirmation();
  testPreviousReceiptSettlementRequiresExplicitStatus();
  testPreviousReceiptVisibilityFollowsReceiptDecision();
  testPreviousReceiptIsHiddenWhenPreviousPreparedMissing();
  testAggregateTemplateUsesExplicitMonths();
  testSelfPaidInvoiceDoesNotRoundManualUnitPrice();
  testReceiptStatusIsOverwrittenInHistory();
  testInsuranceBillingUsesYenRounding();
  testTwoTenthBurdenKeepsYenPrecision();
  testWelfareBillingStillAddsTransport();
  testMassageBillingDoesNotChargeTransport();
  testCarryOverHistoryIsIncluded();
  testBankExportRejectsNullPreparedPayload();
  testBankExportUsesNormalizedBillingJsonWhenMissingFromPayload();
  testBankExportRejectsInvalidBillingJsonShape();
  testBankExportReturnsEmptyWhenNoRows();
  testBankCodesAreNormalizedBeforeValidation();
  testUnifiedNameKeyAcrossBankPatientAndBilling();
  testBankRowsSkipWhenNormalizedLengthsAreInvalid();
  testNameKanaFallsBackToKanjiWhenEmpty();
  testNameKanaIsNormalizedToFullWidth();
  testSkipMessageIncludesReasonsWhenAllRowsInvalid();
  console.log('billingOutput blob guard tests passed');
}

run();
