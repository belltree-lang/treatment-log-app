const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const billingOutputCode = fs.readFileSync(path.join(__dirname, '../src/output/billingOutput.js'), 'utf8');

function createContext(overrides = {}) {
  const ctx = Object.assign({
    console,
    MimeType: {
      GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet',
      MICROSOFT_EXCEL: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      PDF: 'application/pdf'
    }
  }, overrides);
  vm.createContext(ctx);
  vm.runInContext(billingOutputCode, ctx);
  return ctx;
}

const context = createContext();

const {
  calculateInvoiceChargeBreakdown_,
  buildAggregateInvoiceTemplateData_,
  buildBillingInvoiceHtml_,
  buildInvoiceTemplateData_,
  buildInvoiceChargePeriodLabel_
} = context;

if (typeof calculateInvoiceChargeBreakdown_ !== 'function' || typeof buildBillingInvoiceHtml_ !== 'function') {
  throw new Error('Invoice helpers failed to load in the test context');
}

function testInvoiceChargeBreakdown() {
  const breakdown = calculateInvoiceChargeBreakdown_({
    visitCount: 8,
    burdenRate: 1,
    insuranceType: '鍼灸',
    carryOverAmount: 1000
  });

  assert.strictEqual(breakdown.treatmentUnitPrice, 417, '1割の単価が417円に設定される');
  assert.strictEqual(breakdown.treatmentAmount, 3336, '施術料が円単位で計算される');
  assert.strictEqual(breakdown.transportAmount, 264, '交通費は33円×回数で算定される');
  assert.strictEqual(breakdown.grandTotal, 4600, '合計は繰越+施術料+交通費の和となる');
}

function testInvoiceChargeBreakdownUsesCustomTransportPrice() {
  const customContext = createContext({ BILLING_TRANSPORT_UNIT_PRICE: 50 });
  const { calculateInvoiceChargeBreakdown_ } = customContext;

  const breakdown = calculateInvoiceChargeBreakdown_({
    visitCount: 4,
    burdenRate: 2,
    insuranceType: '鍼灸',
    carryOverAmount: 0
  });

  assert.strictEqual(breakdown.transportAmount, 200, '交通費が上書き単価で算出される');
  assert.strictEqual(breakdown.grandTotal, 3536, '合計にも上書き単価の交通費が反映される');
}

function testInvoiceHtmlIncludesBreakdown() {
  const html = buildBillingInvoiceHtml_({
    billingMonth: '202512',
    visitCount: 8,
    burdenRate: 1,
    insuranceType: '鍼灸',
    carryOverAmount: 1000,
    nameKanji: '山田太郎',
    raw: { '住所': '東京都江東区1-2-3' }
  }, '202512');

  assert(html.includes('2025年12月 ご請求書'), '請求月の見出しが含まれる');
  assert(html.includes('前月繰越: 1,000円'), '繰越内訳が含まれる');
  assert(html.includes('施術料（417円 × 8回）'), '施術料の内訳が含まれる');
  assert(html.includes('施術料（417円 × 8回）: 3,336円'), '施術料の金額が含まれる');
  assert(html.includes('交通費（33円 × 8回）'), '交通費の内訳が含まれる');
  assert(html.includes('交通費（33円 × 8回）: 264円'), '交通費の金額が含まれる');
  assert(html.includes('4,600円'), '合計金額がカンマ区切りで表示される');
  assert(html.includes('べるつりー訪問鍼灸マッサージ'), 'タイトルが含まれる');
}

function testInvoiceHtmlEscapesUserInput() {
  const html = buildBillingInvoiceHtml_({
    billingMonth: '202501',
    visitCount: 1,
    burdenRate: 1,
    insuranceType: '鍼灸',
    nameKanji: '<script>alert(1)</script>',
    address: '東京都 <b>江東区</b>'
  }, '202501');

  assert(!html.includes('<script>'), '埋め込みスクリプトはサニタイズされる');
  assert(html.includes('&lt;script&gt;alert(1)&lt;/script&gt;'), '氏名はエスケープされる');
  assert(!html.includes('江東区'), '住所は出力に含まれない');
}

function testAggregateInvoiceHtmlSumsBreakdown() {
  const patientId = 'patient-aggregate';
  const aggregateEntries = {
    '202501': { billingMonth: '202501', patientId, visitCount: 2, burdenRate: 1, insuranceType: '鍼灸' },
    '202502': { billingMonth: '202502', patientId, visitCount: 3, burdenRate: 1, insuranceType: '鍼灸' }
  };
  const aggregateContext = createContext({
    billingNormalizePatientId_: (value) => value,
    getPreparedBillingEntryForMonthCached_: (monthKey, pid) => (pid === patientId ? aggregateEntries[monthKey] : null)
  });
  const { buildBillingInvoiceHtml_ } = aggregateContext;

  const html = buildBillingInvoiceHtml_({
    billingMonth: '202503',
    patientId,
    aggregateTargetMonths: ['202501', '202502'],
    visitCount: 1,
    burdenRate: 1,
    insuranceType: '鍼灸',
    nameKanji: '集計太郎'
  }, '202503');

  assert(html.includes('施術料（417円 × 5回）: 2,085円'), '合算対象月の施術料を合算して表示する');
  assert(html.includes('交通費（33円 × 5回）: 165円'), '合算対象月の交通費を合算して表示する');
  assert(html.includes('2,250円'), '合算対象月の合計を合算結果で表示する');
}

function testInvoiceTemplateRecalculatesSelfPaidBreakdown() {
  const data = buildInvoiceTemplateData_({
    billingMonth: '202501',
    insuranceType: '自費',
    unitPrice: 3333,
    visitCount: 1,
    treatmentAmount: 1000,
    transportAmount: 0,
    grandTotal: 1000
  });

  assert.strictEqual(data.rows[1].amount, 3333, '請求書テンプレートでも手動単価をそのまま乗算する');
  assert.strictEqual(data.rows[2].amount, 33, '自費でも交通費を自動計算する');
  assert.strictEqual(data.grandTotal, 3366, '合計は再計算された施術料と交通費に基づく');
}

function testInvoiceTemplateAddsReceiptDecision() {
  const unpaid = buildInvoiceTemplateData_({
    billingMonth: '202311',
    receiptStatus: 'UNPAID',
    receiptMonths: ['202310'],
    hasPreviousReceiptSheet: true
  });

  assert.strictEqual(unpaid.showReceipt, false, 'UNPAID の月は領収書を表示しない');

  const payable = buildInvoiceTemplateData_({
    billingMonth: '202311',
    receiptStatus: 'PAID',
    receiptMonths: ['202310'],
    hasPreviousReceiptSheet: true
  });
  assert.strictEqual(payable.showReceipt, true, '未回収チェックが無ければ領収書を表示する');
  assert.deepStrictEqual(Array.from(payable.receiptMonths || []), ['202310'], '領収対象月は前月を指す');
  assert.strictEqual(payable.receiptRemark, '', '備考は付与しない');
}

function testAggregateInvoiceTemplateStacksPerMonth() {
  const patientId = 'patient-1';
  const entriesByMonth = {
    '202501': { billingMonth: '202501', patientId, visitCount: 2, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 100 },
    '202502': { billingMonth: '202502', patientId, visitCount: 3, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 0 },
    '202503': { billingMonth: '202503', patientId, visitCount: 4, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 50 }
  };

  const aggregateContext = createContext({
    billingNormalizePatientId_: (value) => value,
    normalizePreparedBilling_: (value) => value,
    loadPreparedBillingWithSheetFallback_: (monthKey) => ({
      billingMonth: monthKey,
      billingJson: entriesByMonth[monthKey] ? [entriesByMonth[monthKey]] : []
    })
  });
  const { buildAggregateInvoiceTemplateData_ } = aggregateContext;

  const data = buildAggregateInvoiceTemplateData_(entriesByMonth['202503'], ['202501', '202502', '202503']);

  assert.deepStrictEqual(Array.from(data.receiptMonths || []), ['202501', '202502'], '合算対象は billingMonth より前の月に限定される');
  assert.strictEqual(data.aggregateMonthTotals.length, 2, '対象月の合計行のみが含まれる');
  assert.strictEqual(data.aggregateRemark, '01月・02月分 施術料金として', '合算対象月の備考が付与される');
  assert.strictEqual(data.chargeMonthLabel, '2025年03月', '請求対象月のラベルは billingMonth に基づく');
}

function testAggregateInvoiceTemplateSummarizesWhenManyMonths() {
  const patientId = 'patient-2';
  const entriesByMonth = {
    '202501': { billingMonth: '202501', patientId, visitCount: 1, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 200 },
    '202502': { billingMonth: '202502', patientId, visitCount: 1, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 0 },
    '202503': { billingMonth: '202503', patientId, visitCount: 1, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 0 },
    '202504': { billingMonth: '202504', patientId, visitCount: 1, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 0 }
  };

  const aggregateContext = createContext({
    billingNormalizePatientId_: (value) => value,
    normalizePreparedBilling_: (value) => value,
    loadPreparedBillingWithSheetFallback_: (monthKey) => ({
      billingMonth: monthKey,
      billingJson: entriesByMonth[monthKey] ? [entriesByMonth[monthKey]] : []
    })
  });
  const { buildAggregateInvoiceTemplateData_ } = aggregateContext;

  const data = buildAggregateInvoiceTemplateData_(entriesByMonth['202504'], ['202501', '202502', '202503', '202504']);

  assert.deepStrictEqual(Array.from(data.receiptMonths || []), ['202501', '202502', '202503'], '請求対象月より前の3ヶ月が合算対象となる');
  assert.strictEqual(data.aggregateMonthTotals.length, 3, '合算対象月の数だけ月別合計行が含まれる');
  assert.strictEqual(data.aggregateRemark, '01月・02月・03月分 施術料金として', '対象月の備考に複数月が含まれる');
}

function testAggregateInvoiceRepresentativeMonthMatchesBillingMonth() {
  const patientId = 'patient-3';
  const entriesByMonth = {
    '202501': { billingMonth: '202501', patientId, visitCount: 1, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 0 },
    '202502': { billingMonth: '202502', patientId, visitCount: 2, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 150 },
    '202503': { billingMonth: '202503', patientId, visitCount: 1, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 300 }
  };

  const aggregateContext = createContext({
    billingNormalizePatientId_: (value) => value,
    normalizePreparedBilling_: (value) => value,
    loadPreparedBillingWithSheetFallback_: (monthKey) => ({
      billingMonth: monthKey,
      billingJson: entriesByMonth[monthKey] ? [entriesByMonth[monthKey]] : []
    })
  });
  const { buildAggregateInvoiceTemplateData_ } = aggregateContext;

  const data = buildAggregateInvoiceTemplateData_(entriesByMonth['202502'], ['202503', '202501', '202502']);

  assert.deepStrictEqual(Array.from(data.receiptMonths || []), ['202501'], 'billingMonth 以前の月のみが合算対象となる');
  assert.strictEqual(data.aggregateMonthTotals.length, 1, '合算対象に含まれた月の数だけ行が生成される');
  assert.strictEqual(data.chargeMonthLabel, '2025年02月', '合算請求の請求対象ラベルは billingMonth に従う');
  assert.strictEqual(data.aggregateRemark, '01月分 施術料金として', '備考には合算対象の月が含まれる');
}

function testAggregateInvoiceFallsBackToLatestMonthWhenBillingMonthAbsent() {
  const patientId = 'patient-4';
  const entriesByMonth = {
    '202501': { billingMonth: '202501', patientId, visitCount: 1, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 400 },
    '202502': { billingMonth: '202502', patientId, visitCount: 2, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 0 },
    '202503': { billingMonth: '202503', patientId, visitCount: 3, burdenRate: 1, insuranceType: '鍼灸', carryOverAmount: 0 }
  };

  const aggregateContext = createContext({
    billingNormalizePatientId_: (value) => value,
    normalizePreparedBilling_: (value) => value,
    loadPreparedBillingWithSheetFallback_: (monthKey) => ({
      billingMonth: monthKey,
      billingJson: entriesByMonth[monthKey] ? [entriesByMonth[monthKey]] : []
    })
  });
  const { buildAggregateInvoiceTemplateData_ } = aggregateContext;

  const data = buildAggregateInvoiceTemplateData_(
    Object.assign({}, entriesByMonth['202501'], { billingMonth: '202412' }),
    ['202501', '202502', '202503']
  );

  assert.deepStrictEqual(Array.from(data.receiptMonths || []), [], 'billingMonth より後の月のみが指定された場合は合算対象が空になる');
  assert.strictEqual(data.aggregateMonthTotals.length, 0, '対象月が無ければ月別合計行も無い');
  assert.strictEqual(data.aggregateRemark, '', '対象月が無い場合は備考を付与しない');
  assert.strictEqual(data.chargeMonthLabel, '2024年12月', 'billingMonth が対象月に含まれなくても請求対象ラベルは billingMonth を用いる');
}

function testInvoiceTemplateDisplaysPeriodLabel() {
  const templateHtml = fs.readFileSync(path.join(__dirname, '../src/invoice_template.html'), 'utf8');
  const labelUsageCount = (templateHtml.match(/charge-period-label/g) || []).length;
  assert(labelUsageCount >= 2, '対象期間ラベルの表示領域が合計近辺に用意されている');

  const singleMonthData = buildInvoiceTemplateData_({
    billingMonth: '202404',
    visitCount: 2,
    burdenRate: 1,
    insuranceType: '鍼灸'
  });
  const singleLabel = buildInvoiceChargePeriodLabel_(singleMonthData);
  assert.strictEqual(singleLabel, '令和6年04月分', '単月の請求では対象期間ラベルを単月表示する');

  const aggregateLabel = buildInvoiceChargePeriodLabel_({
    isAggregateInvoice: true,
    aggregateMonthTotals: [
      { month: '202412' },
      { month: '202501' }
    ]
  });
  assert.strictEqual(aggregateLabel, '令和6年12月分〜令和7年01月分', '年跨ぎの合算は両方に年を含めて範囲表示する');

  const sameYearLabel = buildInvoiceChargePeriodLabel_({
    isAggregateInvoice: true,
    aggregateMonthTotals: [
      { month: '202503' },
      { month: '202505' }
    ]
  });
  assert.strictEqual(sameYearLabel, '令和7年03月分〜05月分', '同一年内の合算は開始月のみ年を含める');
}

function run() {
  testInvoiceChargeBreakdown();
  testInvoiceChargeBreakdownUsesCustomTransportPrice();
  testInvoiceHtmlIncludesBreakdown();
  testInvoiceHtmlEscapesUserInput();
  testAggregateInvoiceHtmlSumsBreakdown();
  testInvoiceTemplateRecalculatesSelfPaidBreakdown();
  testInvoiceTemplateAddsReceiptDecision();
  testAggregateInvoiceTemplateStacksPerMonth();
  testAggregateInvoiceTemplateSummarizesWhenManyMonths();
  testAggregateInvoiceRepresentativeMonthMatchesBillingMonth();
  testAggregateInvoiceFallsBackToLatestMonthWhenBillingMonthAbsent();
  testInvoiceTemplateDisplaysPeriodLabel();
  console.log('billingInvoiceLayout tests passed');
}

run();
