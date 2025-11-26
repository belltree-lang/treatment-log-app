const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const billingLogicCode = fs.readFileSync(path.join(__dirname, '../src/logic/billingLogic.js'), 'utf8');
const context = {};
vm.createContext(context);
vm.runInContext(billingLogicCode, context);

const { calculateBillingAmounts_, normalizeBurdenMultiplier_ } = context;

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

function run() {
  testBurdenRateDigitConversion();
  testMassageBillingExclusion();
  testBillingAmountRoundsToNearestTen();
  console.log('billingLogic tests passed');
}

run();
