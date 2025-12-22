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

const { calculateInvoiceChargeBreakdown_, buildBillingInvoiceHtml_, buildInvoiceTemplateData_ } = context;

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
    receiptStatus: 'UNPAID'
  });

  assert.strictEqual(unpaid.showReceipt, false, 'UNPAID の月は領収書を表示しない');

  const payable = buildInvoiceTemplateData_({ billingMonth: '202311', receiptStatus: 'PAID' });
  assert.strictEqual(payable.showReceipt, true, '未回収チェックが無ければ領収書を表示する');
  assert.deepStrictEqual(Array.from(payable.receiptMonths || []), ['202310'], '領収対象月は前月を指す');
  assert.strictEqual(payable.receiptRemark, '', '備考は付与しない');
}

function run() {
  testInvoiceChargeBreakdown();
  testInvoiceChargeBreakdownUsesCustomTransportPrice();
  testInvoiceHtmlIncludesBreakdown();
  testInvoiceHtmlEscapesUserInput();
  testInvoiceTemplateRecalculatesSelfPaidBreakdown();
  testInvoiceTemplateAddsReceiptDecision();
  console.log('billingInvoiceLayout tests passed');
}

run();
