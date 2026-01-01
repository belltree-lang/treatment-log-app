// Phase2仕様で暗黙的な自動合算（scheduled / 金額起点 / bankFlags 自動合算）を廃止したため、関連テストは削除済み。
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const mainCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'main.gs'), 'utf8');

function createContext(preparedByMonth) {
  const ctx = {
    console: { log: () => {}, warn: () => {} },
    preparedByMonth: preparedByMonth || {},
    normalizeMoneyNumber_: value => Number(value) || 0,
    billingLogger_: { log: () => {} }
  };
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
  ctx.loadPreparedBillingWithSheetFallback_ = monthKey => ctx.preparedByMonth[monthKey] || null;
  ctx.buildReceiptMonthsFromBankUnpaid_ = (patientId, anchorMonth) => {
    const prev = ctx.resolvePreviousBillingMonthKey_(anchorMonth);
    return prev ? [prev, anchorMonth] : [anchorMonth];
  };

  return ctx;
}

(function testAggregateInvoiceCombinesUnpaidMonthsWhenAggregateMonthSpecified() {
  const preparedByMonth = {
    '202401': {
      billingMonth: '202401',
      billingJson: [{ patientId: 'P01', grandTotal: 1000 }],
      bankFlagsByPatient: { P01: { ae: true, af: false } }
    },
    '202402': {
      billingMonth: '202402',
      billingJson: [{ patientId: 'P01', grandTotal: 2000, aggregateUntilMonth: '202402' }],
      bankFlagsByPatient: { P01: { ae: true, af: false } }
    }
  };

  const context = createContext(preparedByMonth);
  const result = context.applyAggregateInvoiceRulesFromBankFlags_(preparedByMonth['202402']);
  const entry = result.billingJson[0];

  assert.strictEqual(entry.aggregateStatus, 'confirmed');
  assert.strictEqual(entry.grandTotal, 3000);
  assert.deepStrictEqual([].concat(entry.aggregateTargetMonths || []), ['202401', '202402']);
  assert.strictEqual(entry.skipReceipt, true);
})();

(function testUnpaidWithoutAggregateMonthDoesNotAggregate() {
  const preparedByMonth = {
    '202403': {
      billingMonth: '202403',
      billingJson: [{ patientId: 'P02', grandTotal: 1500 }],
      bankFlagsByPatient: { P02: { ae: true, af: false } }
    }
  };

  const context = createContext(preparedByMonth);
  const result = context.applyAggregateInvoiceRulesFromBankFlags_(preparedByMonth['202403']);
  const entry = result.billingJson[0];

  assert.strictEqual(entry.aggregateStatus, undefined);
  assert.strictEqual(entry.grandTotal, 1500);
  assert.strictEqual(entry.skipReceipt, undefined);
})();  
