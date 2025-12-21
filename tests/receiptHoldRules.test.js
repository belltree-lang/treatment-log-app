const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const mainCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'main.gs'), 'utf8');

function createContext(overrides = {}) {
  const ctx = Object.assign({ console }, overrides);
  vm.createContext(ctx);
  vm.runInContext(mainCode, ctx);
  return ctx;
}

(function testUnpaidChecksOnlyHoldMatchingPatients() {
  const ctx = createContext();
  ctx.summarizeBankWithdrawalSheet_ = () => ({ billingMonth: '202501', unpaidChecked: 2 });
  ctx.resolveUnpaidCheckedPatientIds_ = () => new Set(['P-HOLD']);

  const prepared = {
    billingMonth: '202501',
    receiptStatus: 'AGGREGATE',
    aggregateUntilMonth: '202504',
    billingJson: [
      { patientId: 'P-HOLD', billingMonth: '202501' },
      { patientId: 'P-OK', billingMonth: '202501' }
    ]
  };

  const result = ctx.applyReceiptRulesFromUnpaidCheck_(prepared);
  const holdRow = result.billingJson.find(row => row.patientId === 'P-HOLD');
  const okRow = result.billingJson.find(row => row.patientId === 'P-OK');

  assert.strictEqual(result.receiptStatus, 'AGGREGATE', '集計設定は全体の設定を保持する');
  assert.strictEqual(result.aggregateUntilMonth, '202504', '全体の合算終了月を保持する');
  assert.strictEqual(holdRow.receiptStatus, 'HOLD', '未回収チェックの患者のみ HOLD になる');
  assert.strictEqual(holdRow.aggregateUntilMonth, '', 'HOLD の患者には合算終了月を付与しない');
  assert.strictEqual(okRow.receiptStatus, 'AGGREGATE', '未回収チェック無しの患者は合算設定を維持する');
  assert.strictEqual(okRow.aggregateUntilMonth, '202504', '未回収チェック無しの患者は合算終了月を維持する');
})();

console.log('receipt hold rules tests passed');
