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
  ctx.preparedByMonth = {};
  ctx.getPreparedBillingForMonthCached_ = monthKey => ctx.preparedByMonth[monthKey] || null;
  ctx.getPreparedBillingEntryForMonthCached_ = (monthKey, patientId) => {
    const payload = ctx.preparedByMonth[monthKey];
    if (!payload || !Array.isArray(payload.billingJson)) return null;
    return payload.billingJson.find(row => ctx.billingNormalizePatientId_(row && row.patientId) === ctx.billingNormalizePatientId_(patientId)) || null;
  };
  ctx.isPatientCheckedUnpaidInBankWithdrawalSheet_ = () => false;
  return ctx;
}

(function testAggregateReceiptGeneratedFromUnpaidHistory() {
  const context = createContext();

  context.preparedByMonth['202410'] = {
    billingMonth: '202410',
    billingJson: [{ patientId: 'P01' }],
    bankFlagsByPatient: { P01: { ae: true } }
  };
  context.preparedByMonth['202411'] = {
    billingMonth: '202411',
    billingJson: [{ patientId: 'P01' }],
    bankFlagsByPatient: { P01: { ae: true } }
  };
  context.buildReceiptMonthBreakdownForEntry_ = (patientId, months) => months.map((month, idx) => ({
    month,
    amount: (idx + 1) * 1000
  }));

  const prepared = {
    billingMonth: '202412',
    billingJson: [
      { patientId: 'P01' },
      { patientId: 'P02' }
    ],
    bankFlagsByPatient: { P01: { ae: false }, P02: { ae: false } }
  };

  const enriched = context.attachPreviousReceiptAmounts_(prepared);
  const patientWithUnpaid = enriched.billingJson[0];
  const patientWithoutUnpaid = enriched.billingJson[1];

  assert.deepStrictEqual(
    Array.from(patientWithUnpaid.receiptMonths || []),
    ['202410', '202411', '202412'],
    'aggregate receipt includes consecutive unpaid history and current month'
  );
  assert.ok(
    Array.isArray(patientWithUnpaid.receiptMonthBreakdown) && patientWithUnpaid.receiptMonthBreakdown.length === 3,
    'aggregate receipt keeps month breakdown'
  );
  assert.deepStrictEqual(
    Array.from(patientWithoutUnpaid.receiptMonths || []),
    ['202411'],
    'patients without aggregate invoices keep the default single-month receipt'
  );
})();

console.log('receipt aggregation from bank sheet tests passed');
