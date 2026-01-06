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
  ctx.getPreparedBillingForMonthCached_ = monthKey => ctx.preparedByMonth[monthKey] || null;

  return ctx;
}

(function testAggregateInvoiceCombinesUnpaidHistoryWhenPaymentArrives() {
  const preparedByMonth = {
    '202401': {
      billingMonth: '202401',
      billingJson: [{ patientId: 'P01', grandTotal: 1000 }],
      totalsByPatient: { P01: { grandTotal: 1000 } },
      bankFlagsByPatient: { P01: { ae: true, af: false } }
    },
    '202402': {
      billingMonth: '202402',
      billingJson: [{ patientId: 'P01', grandTotal: 2000 }],
      totalsByPatient: { P01: { grandTotal: 2000 } },
      bankFlagsByPatient: { P01: { ae: false, af: false } }
    }
  };

  const context = createContext(preparedByMonth);
  const result = context.applyAggregateInvoiceRulesFromBankFlags_(preparedByMonth['202402']);
  const entry = result.billingJson[0];

  assert.strictEqual(entry.aggregateStatus, 'confirmed');
  assert.strictEqual(entry.grandTotal, 3000);
  assert.deepStrictEqual([].concat(entry.aggregateTargetMonths || []), ['202401', '202402']);
  assert.strictEqual(entry.skipReceipt, false);
})();

(function testUnpaidWithoutAggregateMonthDoesNotAggregate() {
  const preparedByMonth = {
    '202403': {
      billingMonth: '202403',
      billingJson: [{ patientId: 'P02', grandTotal: 1500 }],
      totalsByPatient: { P02: { grandTotal: 1500 } },
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

(function testConsecutiveUnpaidAggregatedOnceOnRecovery() {
  const preparedByMonth = {
    '202401': {
      billingMonth: '202401',
      billingJson: [{ patientId: 'P03', grandTotal: 1200 }],
      totalsByPatient: { P03: { grandTotal: 1200 } },
      bankFlagsByPatient: { P03: { ae: true, af: false } }
    },
    '202402': {
      billingMonth: '202402',
      billingJson: [{ patientId: 'P03', grandTotal: 1300 }],
      totalsByPatient: { P03: { grandTotal: 1300 } },
      bankFlagsByPatient: { P03: { ae: true, af: false } }
    },
    '202403': {
      billingMonth: '202403',
      billingJson: [{ patientId: 'P03', grandTotal: 1400 }],
      totalsByPatient: { P03: { grandTotal: 1400 } },
      bankFlagsByPatient: { P03: { ae: false, af: false } }
    }
  };

  const context = createContext(preparedByMonth);

  const duringUnpaid = context.applyAggregateInvoiceRulesFromBankFlags_(preparedByMonth['202402']);
  const duringEntry = duringUnpaid.billingJson[0];
  assert.strictEqual(duringEntry.aggregateStatus, undefined, 'no aggregation while unpaid continues');
  assert.strictEqual(duringEntry.grandTotal, 1300, 'amount stays untouched during unpaid');

  const recovered = context.applyAggregateInvoiceRulesFromBankFlags_(preparedByMonth['202403']);
  const recoveredEntry = recovered.billingJson[0];
  assert.strictEqual(recoveredEntry.aggregateStatus, 'confirmed', 'aggregates once when payment resumes');
  assert.deepStrictEqual(
    [].concat(recoveredEntry.aggregateTargetMonths || []),
    ['202401', '202402', '202403'],
    'includes consecutive unpaid months plus current month'
  );
  assert.strictEqual(recoveredEntry.grandTotal, 3900, 'sums consecutive unpaid months and current month');
  assert.strictEqual(recoveredEntry.skipReceipt, false, 'receipt remains eligible on aggregated settlement');
})();

(function testAggregationDoesNotRepeatAfterRecoveryMonth() {
  const preparedByMonth = {
    '202401': {
      billingMonth: '202401',
      billingJson: [{ patientId: 'P04', grandTotal: 1000 }],
      totalsByPatient: { P04: { grandTotal: 1000 } },
      bankFlagsByPatient: { P04: { ae: true } }
    },
    '202402': {
      billingMonth: '202402',
      billingJson: [{ patientId: 'P04', grandTotal: 1100 }],
      totalsByPatient: { P04: { grandTotal: 1100 } },
      bankFlagsByPatient: { P04: { ae: true } }
    },
    '202403': {
      billingMonth: '202403',
      billingJson: [{ patientId: 'P04', grandTotal: 1200 }],
      totalsByPatient: { P04: { grandTotal: 1200 } },
      bankFlagsByPatient: { P04: { ae: false } }
    },
    '202404': {
      billingMonth: '202404',
      billingJson: [{ patientId: 'P04', grandTotal: 1300 }],
      totalsByPatient: { P04: { grandTotal: 1300 } },
      bankFlagsByPatient: { P04: { ae: false } }
    }
  };

  const context = createContext(preparedByMonth);
  const recoveryMonthResult = context.applyAggregateInvoiceRulesFromBankFlags_(preparedByMonth['202403']);
  const recoveryEntry = recoveryMonthResult.billingJson[0];
  assert.strictEqual(recoveryEntry.aggregateStatus, 'confirmed', 'aggregates once at recovery month');
  assert.deepStrictEqual(
    [].concat(recoveryEntry.aggregateTargetMonths || []),
    ['202401', '202402', '202403'],
    'aggregates consecutive unpaid months and settlement month'
  );

  // Simulate persisting the aggregated month before the next run
  preparedByMonth['202403'] = recoveryMonthResult;

  const nextMonthResult = context.applyAggregateInvoiceRulesFromBankFlags_(preparedByMonth['202404']);
  const nextEntry = nextMonthResult.billingJson[0];
  assert.strictEqual(nextEntry.aggregateStatus, undefined, 'does not aggregate again after recovery month');
  assert.deepStrictEqual([].concat(nextEntry.aggregateTargetMonths || []), [], 'no aggregate months on the following month');
})();

(function testNewUnpaidCycleCanAggregateAfterPriorRecoveryInCache() {
  const preparedByMonth = {
    // First unpaid cycle aggregated at 202402 and cached.
    '202401': {
      billingMonth: '202401',
      billingJson: [{ patientId: 'P05', grandTotal: 900 }],
      totalsByPatient: { P05: { grandTotal: 900 } },
      bankFlagsByPatient: { P05: { ae: true } }
    },
    '202402': {
      billingMonth: '202402',
      billingJson: [{
        patientId: 'P05',
        grandTotal: 1800,
        aggregateTargetMonths: ['202401', '202402'],
        aggregateStatus: 'confirmed'
      }],
      totalsByPatient: { P05: { grandTotal: 1800 } },
      bankFlagsByPatient: { P05: { ae: false } }
    },
    // Second unpaid cycle begins later.
    '202405': {
      billingMonth: '202405',
      billingJson: [{ patientId: 'P05', grandTotal: 1000 }],
      totalsByPatient: { P05: { grandTotal: 1000 } },
      bankFlagsByPatient: { P05: { ae: true } }
    },
    '202406': {
      billingMonth: '202406',
      billingJson: [{ patientId: 'P05', grandTotal: 1100 }],
      totalsByPatient: { P05: { grandTotal: 1100 } },
      bankFlagsByPatient: { P05: { ae: false } }
    }
  };

  const context = createContext(preparedByMonth);
  const monthCache = { preparedByMonth: preparedByMonth };
  const result = context.applyAggregateInvoiceRulesFromBankFlags_(preparedByMonth['202406'], monthCache);
  const entry = result.billingJson[0];

  assert.strictEqual(entry.aggregateStatus, 'confirmed', 'allows aggregation for a new unpaid cycle');
  assert.deepStrictEqual(
    [].concat(entry.aggregateTargetMonths || []),
    ['202405', '202406'],
    'aggregates only the latest unpaid streak and recovery month'
  );
})();  
