const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const mainCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'main.gs'), 'utf8');

function createContext() {
  const ctx = { console: { warn: () => {}, log: () => {} } };
  vm.createContext(ctx);
  vm.runInContext(mainCode, ctx);

  ctx.normalizeBillingMonthKeySafe_ = value => String(value || '').trim();
  ctx.normalizeBillingMonthInput = value => ({
    key: String(value || ''),
    year: Number(String(value || '').slice(0, 4)) || 2024,
    month: Number(String(value || '').slice(4, 6)) || 1
  });
  ctx.resolvePreviousBillingMonthKey_ = billingMonth => {
    const normalized = ctx.normalizeBillingMonthKeySafe_(billingMonth);
    if (!normalized) return '';
    const year = Number(normalized.slice(0, 4));
    const month = Number(normalized.slice(4, 6));
    const prevMonth = month === 1 ? 12 : month - 1;
    const prevYear = month === 1 ? year - 1 : year;
    return String(prevYear).padStart(4, '0') + String(prevMonth).padStart(2, '0');
  };
  ctx.billingNormalizePatientId_ = pid => (pid ? String(pid).trim() : '');

  return ctx;
}

(function testReceiptMonthsBackfillFromBankSheet() {
  const context = createContext();

  const bankAmounts = {
    '202411': { P01: 1200 },
    '202410': { P01: 800 }
  };

  context.collectPreviousReceiptAmountsFromBankSheet_ = monthKey => ({
    hasSheet: true,
    amounts: bankAmounts[monthKey] || {}
  });
  context.buildReceiptMonthsFromBankUnpaid_ = patientId => (patientId === 'P01'
    ? ['202410', '202411']
    : ['202411']
  );
  context.collectBankWithdrawalAmountsByPatient_ = monthKey => bankAmounts[monthKey] || {};

  const prepared = {
    billingMonth: '202412',
    billingJson: [
      { patientId: 'P01' },
      { patientId: 'P02' }
    ]
  };

  const enriched = context.attachPreviousReceiptAmounts_(prepared);
  const patientWithUnpaid = enriched.billingJson[0];
  const patientWithoutUnpaid = enriched.billingJson[1];

  assert.deepStrictEqual(
    patientWithUnpaid.receiptMonths,
    ['202410', '202411'],
    'receipt months include unpaid history'
  );
  assert.strictEqual(
    patientWithUnpaid.previousReceiptAmount,
    2000,
    'previousReceiptAmount aggregates all receipt months'
  );
  assert.deepStrictEqual(
    patientWithoutUnpaid.receiptMonths,
    ['202411'],
    'patients without unpaid history only include anchor month'
  );
  assert.strictEqual(
    patientWithoutUnpaid.previousReceiptAmount,
    0,
    'patients without amounts keep zero aggregate'
  );
})();

console.log('receipt aggregation from bank sheet tests passed');
