const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const billingLogicCode = fs.readFileSync(path.join(__dirname, '../src/logic/billingLogic.js'), 'utf8');

function createLogicContext(overrides = {}) {
  const ctx = Object.assign({
    normalizeBillingNameKey_: value => String(value || '').trim()
  }, overrides);
  vm.createContext(ctx);
  vm.runInContext(billingLogicCode, ctx);
  return ctx;
}

const context = createLogicContext();

const { calculateBillingAmounts_, normalizeBurdenMultiplier_ } = context;
const { generateBillingJsonFromSource } = context;

if (typeof calculateBillingAmounts_ !== 'function' || typeof normalizeBurdenMultiplier_ !== 'function') {
  throw new Error('Billing logic functions failed to load in the test context');
}

function testBurdenRateDigitConversion() {
  assert.strictEqual(normalizeBurdenMultiplier_(1, ''), 0.1, '1割は0.1に換算される');
  assert.strictEqual(normalizeBurdenMultiplier_(3, ''), 0.3, '3割は0.3に換算される');
  assert.strictEqual(normalizeBurdenMultiplier_(0.3, ''), 0.3, '小数表記はそのまま利用される');
  assert.strictEqual(normalizeBurdenMultiplier_(2, '自費'), 1, '自費は常に1倍');
}

function testMassageBillingExclusion() {
  const result = calculateBillingAmounts_({
    visitCount: 4,
    insuranceType: 'マッサージ',
    burdenRate: 3,
    unitPrice: 5000,
    carryOverAmount: 2000
  });

  assert.strictEqual(result.unitPrice, 0, 'マッサージは単価を計上しない');
  assert.strictEqual(result.total, 0, 'マッサージは合計0円');
  assert.strictEqual(result.billingAmount, 0, 'マッサージは請求額0円');
  assert.strictEqual(result.grandTotal, 2000, '繰越額のみが合計に残る');
}

function testBillingAmountRoundsToNearestTen() {
  const result = calculateBillingAmounts_({
    visitCount: 6,
    insuranceType: '鍼灸',
    burdenRate: 3,
    unitPrice: 4170,
    carryOverAmount: 0
  });

  assert.strictEqual(result.visits, 6, '施術回数が正しく反映される');
  assert.strictEqual(result.unitPrice, 4170, '単価がデフォルト料金で設定される');
  assert.strictEqual(result.billingAmount, 7510, '請求額は10円単位に四捨五入される');
}

function testPaidStatusIsIncludedInBillingJson() {
  const source = {
    billingMonth: '202501',
    patients: { '001': { nameKanji: '山田太郎', burdenRate: 1, insuranceType: '鍼灸', unitPrice: 1000 } },
    treatmentVisitCounts: { '001': 2 },
    bankStatuses: { '001': { bankStatus: 'OK', paidStatus: '回収' } },
    bankInfoByName: {}
  };

  const billingJson = generateBillingJsonFromSource(source);
  assert.strictEqual(billingJson[0].paidStatus, '回収', 'BillingJson に領収状態が含まれる');
  assert.strictEqual(billingJson[0].bankStatus, 'OK', '従来の入金ステータスも維持される');
}

function testCarryOverIncludesUnpaidHistory() {
  const source = {
    billingMonth: '202503',
    patients: {
      '010': {
        nameKanji: 'テスト太郎',
        burdenRate: 3,
        insuranceType: '鍼灸',
        carryOverAmount: 500
      }
    },
    treatmentVisitCounts: { '010': 2 },
    carryOverByPatient: { '010': 1500 }
  };

  const billingJson = generateBillingJsonFromSource(source);
  assert.strictEqual(billingJson[0].carryOverAmount, 2000, '患者シートの繰越と未回収が合算される');
  assert.strictEqual(billingJson[0].carryOverFromHistory, 1500, '未回収分が別途保持される');
  assert.strictEqual(billingJson[0].grandTotal, 4566, '合計には繰越を含めた金額が反映される');
}

function testCustomTransportUnitPriceIsUsed() {
  const customContext = createLogicContext({ BILLING_TRANSPORT_UNIT_PRICE: 50 });
  const { calculateBillingAmounts_ } = customContext;

  const result = calculateBillingAmounts_({
    visitCount: 2,
    insuranceType: '鍼灸',
    burdenRate: 3
  });

  assert.strictEqual(result.transportAmount, 100, '交通費が上書き単価で計算される');
  assert.strictEqual(result.grandTotal, 2600, '合計も上書き単価を反映する');
}

function testFullWidthNumbersAreParsed() {
  const result = calculateBillingAmounts_({
    visitCount: '４',
    insuranceType: '自費',
    burdenRate: '１',
    unitPrice: '４,１７０',
    carryOverAmount: '１，５００'
  });

  assert.strictEqual(result.visits, 4, '全角の回数も集計対象になる');
  assert.strictEqual(result.unitPrice, 4170, '全角の単価も正しく解釈される');
  assert.strictEqual(result.carryOverAmount, 1500, '全角の繰越額も数値化される');
  assert.strictEqual(result.grandTotal, 18312, '全角入力でも合計が正しく算出される');
}

function run() {
  testBurdenRateDigitConversion();
  testMassageBillingExclusion();
  testBillingAmountRoundsToNearestTen();
  testPaidStatusIsIncludedInBillingJson();
  testCarryOverIncludesUnpaidHistory();
  testCustomTransportUnitPriceIsUsed();
  testFullWidthNumbersAreParsed();
  console.log('billingLogic tests passed');
}

run();
