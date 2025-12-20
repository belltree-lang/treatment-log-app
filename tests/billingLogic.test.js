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

const { calculateBillingAmounts_, normalizeBurdenMultiplier_, resolveInvoiceUnitPrice_ } = context;
const { generateBillingJsonFromSource, normalizeMedicalAssistanceFlag_ } = context;

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

function testResponsibleNameUsesLatestVisitOnly() {
  const source = {
    billingMonth: '202512',
    patients: {
      '999': {
        nameKanji: '担当確認',
        burdenRate: 1,
        insuranceType: '鍼灸',
        unitPrice: 1000
      }
    },
    treatmentVisitCounts: { '999': 2 },
    staffByPatient: { '999': ['late@example.com', 'early@example.com'] },
    staffDirectory: {
      'late@example.com': '最終担当',
      'early@example.com': '過去担当'
    },
    staffDisplayByPatient: { '999': ['最終担当', '過去担当'] },
    bankStatuses: {}
  };

  const billingJson = generateBillingJsonFromSource(source);
  assert.strictEqual(billingJson[0].responsibleName, '最終担当', '請求書フォルダの担当者は最終訪問スタッフのみ');
  assert.deepStrictEqual(billingJson[0].responsibleNames, ['最終担当', '過去担当'], '担当履歴は配列で保持される');
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

function testMedicalSubsidyExcludesBillingEntries() {
  const logs = [];
  const contextWithLogger = createLogicContext({
    billingLogger_: { log: message => logs.push(message) }
  });
  const { generateBillingJsonFromSource } = contextWithLogger;

  const source = {
    billingMonth: '202501',
    patients: {
      '148': {
        nameKanji: '磯部登志子',
        burdenRate: 1,
        insuranceType: '鍼灸',
        unitPrice: 1000,
        medicalSubsidy: 1
      }
    },
    treatmentVisitCounts: { '148': 2 },
    bankStatuses: {}
  };

  const billingJson = generateBillingJsonFromSource(source);
  assert.strictEqual(billingJson.length, 0, '医療助成対象は請求一覧から除外される');
  assert.ok(logs.some(msg => msg.includes('患者ID 148') && msg.includes('請求対象外')), '除外ログが出力される');
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

function testMedicalAssistanceNormalizationIsStrict() {
  assert.strictEqual(normalizeMedicalAssistanceFlag_(1), 1, '数値1は有効な医療助成フラグ');
  assert.strictEqual(normalizeMedicalAssistanceFlag_('1'), 1, '文字列1も有効な医療助成フラグ');
  assert.strictEqual(normalizeMedicalAssistanceFlag_(true), 1, 'trueは医療助成ありとみなす');
  assert.strictEqual(normalizeMedicalAssistanceFlag_(0), 0, '数値0は助成なし');
  assert.strictEqual(normalizeMedicalAssistanceFlag_(null), 0, 'nullは助成なし');
  assert.strictEqual(normalizeMedicalAssistanceFlag_('yes'), 0, '指定以外の文字列は助成なし');
}

function testBillingJsonIncludesInsuranceMeta() {
  const source = {
    billingMonth: '202504',
    patients: {
      '021': {
        nameKanji: '保険検証',
        burdenRate: 3,
        insuranceType: '鍼灸',
        unitPrice: 5500,
        medicalAssistance: '1'
      }
    },
    treatmentVisitCounts: { '021': 2 },
    bankStatuses: {}
  };

  const billingJson = generateBillingJsonFromSource(source);
  const entry = billingJson[0];

  assert.strictEqual(entry.insuranceType, '鍼灸', '保険種別が必ず含まれる');
  assert.strictEqual(entry.burdenRate, 3, '負担割合は数値で正規化される');
  assert.strictEqual(entry.medicalAssistance, 1, '医療助成は数値フラグで保持される');
  assert.strictEqual(entry.manualUnitPrice, 5500, '手動入力の単価が保持される');
  assert.strictEqual(entry.unitPrice, 5500, 'resolveInvoiceUnitPrice_へ渡した値が反映される');
}

function testFullWidthNumbersAreParsedAndAppliedForSelfPaidManualPrice() {
  const result = calculateBillingAmounts_({
    visitCount: '４',
    insuranceType: '自費',
    burdenRate: '１',
    unitPrice: '４,１７０',
    carryOverAmount: '１，５００'
  });

  assert.strictEqual(result.visits, 4, '全角の回数も集計対象になる');
  assert.strictEqual(result.unitPrice, 4170, '自費でも手動入力の単価が優先される');
  assert.strictEqual(result.treatmentAmount, 16680, '自費で単価が指定された場合は施術料を計上する');
  assert.strictEqual(result.transportAmount, 132, '自費で単価を入力した場合でも交通費を計上する');
  assert.strictEqual(result.carryOverAmount, 1500, '全角の繰越額も数値化される');
  assert.strictEqual(result.grandTotal, 18312, '手動単価と繰越・交通費を合算した金額になる');
}

function testSelfPaidDefaultsToZeroWithoutManualUnitPrice() {
  const result = calculateBillingAmounts_({
    visitCount: 3,
    insuranceType: '自費',
    burdenRate: 2,
    unitPrice: '',
    carryOverAmount: 500
  });

  assert.strictEqual(result.unitPrice, 0, '単価未設定の自費は0円のまま');
  assert.strictEqual(result.treatmentAmount, 0, '単価未設定なら施術料は計上しない');
  assert.strictEqual(result.transportAmount, 0, '単価未設定なら交通費も計上しない');
  assert.strictEqual(result.grandTotal, 500, '繰越額のみが合計に反映される');
}

function testSelfPaidManualPriceIsNotRounded() {
  const result = calculateBillingAmounts_({
    visitCount: 1,
    insuranceType: '自費',
    unitPrice: 3333,
    carryOverAmount: 0
  });

  assert.strictEqual(result.treatmentAmount, 3333, '自費の手動単価はそのまま乗算される');
  assert.strictEqual(result.billingAmount, 3333, '自費では10円単位に丸めず請求する');
  assert.strictEqual(result.grandTotal, 3366, '施術料と交通費を合算した金額が保持される');
}

function testInvoiceUnitPriceResolutionPriority() {
  const manualOverridesZero = resolveInvoiceUnitPrice_('生保', 3, 5000, 1);
  assert.strictEqual(manualOverridesZero, 5000, '手動単価は生保や医療助成より優先される');

  const medicalAssistanceZero = calculateBillingAmounts_({
    visitCount: 3,
    insuranceType: '鍼灸',
    burdenRate: 3,
    unitPrice: '',
    medicalAssistance: 1
  });
  assert.strictEqual(medicalAssistanceZero.unitPrice, 0, '医療助成では単価が0円になる');
  assert.strictEqual(medicalAssistanceZero.treatmentAmount, 0, '医療助成では施術料を計上しない');

  const lifeProtectionZero = calculateBillingAmounts_({
    visitCount: 2,
    insuranceType: '生活扶助',
    burdenRate: 2,
    unitPrice: null
  });
  assert.strictEqual(lifeProtectionZero.unitPrice, 0, '生活扶助も生保扱いで単価0円になる');
  assert.strictEqual(lifeProtectionZero.treatmentAmount, 0, '生保扱いでは施術料を計上しない');

  const selfPaidByBurdenRate = calculateBillingAmounts_({
    visitCount: 1,
    insuranceType: '鍼灸',
    burdenRate: '自費',
    unitPrice: 4000
  });
  assert.strictEqual(selfPaidByBurdenRate.unitPrice, 4000, '自費負担割合でも手動単価を利用する');
  assert.strictEqual(selfPaidByBurdenRate.billingAmount, 4000, '自費負担割合では丸めずに請求する');
  assert.strictEqual(selfPaidByBurdenRate.transportAmount, 33, '自費でも交通費を計上する');
}

function run() {
  testBurdenRateDigitConversion();
  testMassageBillingExclusion();
  testBillingAmountRoundsToNearestTen();
  testPaidStatusIsIncludedInBillingJson();
  testResponsibleNameUsesLatestVisitOnly();
  testCarryOverIncludesUnpaidHistory();
  testMedicalSubsidyExcludesBillingEntries();
  testCustomTransportUnitPriceIsUsed();
  testMedicalAssistanceNormalizationIsStrict();
  testFullWidthNumbersAreParsedAndAppliedForSelfPaidManualPrice();
  testSelfPaidDefaultsToZeroWithoutManualUnitPrice();
  testSelfPaidManualPriceIsNotRounded();
  testInvoiceUnitPriceResolutionPriority();
  testBillingJsonIncludesInsuranceMeta();
  console.log('billingLogic tests passed');
}

run();
