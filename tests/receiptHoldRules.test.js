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
  assert.strictEqual(holdRow.unpaidChecked, true, '未回収チェックの患者のみ unpaidChecked が付与される');
  assert.strictEqual(okRow.unpaidChecked, false, '未回収チェック無しの患者には unpaidChecked が付与されない');
  assert.strictEqual(holdRow.receiptStatus, undefined, '行の領収書状態は変更しない');
  assert.strictEqual(okRow.receiptStatus, undefined, '行の領収書状態は変更しない');
})();

console.log('receipt hold rules tests passed');
