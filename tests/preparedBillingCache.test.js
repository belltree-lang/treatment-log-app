const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const mainCode = fs.readFileSync(path.join(__dirname, '../src/main.gs'), 'utf8');

function buildValidPreparedPayload(ctx, overrides = {}) {
  return Object.assign({
    schemaVersion: ctx.PREPARED_BILLING_SCHEMA_VERSION || 2,
    billingMonth: '202501',
    billingJson: [],
    carryOverLedger: [],
    carryOverLedgerMeta: {},
    carryOverLedgerByPatient: {},
    unpaidHistory: [],
    visitsByPatient: {},
    totalsByPatient: {},
    patients: {},
    bankInfoByName: {},
    staffByPatient: {},
    staffDirectory: {},
    staffDisplayByPatient: {},
    billingOverrideFlags: {},
    carryOverByPatient: {},
    bankAccountInfoByPatient: {}
  }, overrides);
}

function createMainContext(overrides = {}) {
  const baseCache = {
    getScriptCache: () => ({
      get: () => null,
      remove: () => {}
    })
  };
  const ctx = Object.assign({ console, CacheService: baseCache }, overrides);
  vm.createContext(ctx);
  vm.runInContext(mainCode, ctx);
  return Object.assign(ctx, overrides);
}

function testValidationRequiresLedgerFields() {
  const ctx = createMainContext();
  const { validatePreparedBillingPayload_ } = ctx;
  const payload = buildValidPreparedPayload(ctx, { carryOverLedger: undefined });

  const result = validatePreparedBillingPayload_(payload, '202501');
  assert.strictEqual(result.ok, false, 'ledger不足のpayloadはrejectされる');
  assert.strictEqual(result.reason, 'carryOverLedger missing');
}

function testValidationAcceptsCompletePayload() {
  const ctx = createMainContext();
  const { validatePreparedBillingPayload_ } = ctx;
  const payload = buildValidPreparedPayload(ctx);

  const result = validatePreparedBillingPayload_(payload, '202501');
  assert.strictEqual(result.ok, true, '必須フィールドが揃えば検証OK');
  assert.strictEqual(result.billingMonth, '202501');
}

function testSchemaVersionIsValidated() {
  const ctx = createMainContext();
  const { validatePreparedBillingPayload_ } = ctx;
  const payload = buildValidPreparedPayload(ctx, { schemaVersion: 1 });

  const result = validatePreparedBillingPayload_(payload, '202501');
  assert.strictEqual(result.ok, false, '古いスキーマバージョンはrejectされる');
  assert.strictEqual(result.reason, 'schemaVersion mismatch');
}

function testBankExportRejectsNonArrayBillingJson() {
  const ctx = createMainContext({
    normalizeBillingMonthInput: input => ({ key: String(input) }),
    loadPreparedBilling_: () => ({ billingJson: {} }),
    normalizePreparedBilling_: payload => payload,
    exportBankTransferDataForPrepared_: () => assert.fail('invalid payload should not reach exporter')
  });

  assert.throws(() => {
    ctx.generateBankTransferDataFromCache('202501');
  }, /CarryOverLedger シート/);
}

function testBankExportPassesWhenArrayProvided() {
  const exportCalls = [];
  const ctx = createMainContext({
    normalizeBillingMonthInput: input => ({ key: String(input) }),
    loadPreparedBilling_: () => ({ billingJson: [], billingMonth: '202501' }),
    normalizePreparedBilling_: payload => payload,
    exportBankTransferDataForPrepared_: prepared => {
      exportCalls.push(prepared.billingMonth);
      return { billingMonth: prepared.billingMonth, inserted: 1 };
    }
  });

  const result = ctx.generateBankTransferDataFromCache('202501');
  assert.deepStrictEqual(exportCalls, ['202501']);
  assert.strictEqual(result.inserted, 1);
}

function run() {
  testValidationRequiresLedgerFields();
  testValidationAcceptsCompletePayload();
  testSchemaVersionIsValidated();
  testBankExportRejectsNonArrayBillingJson();
  testBankExportPassesWhenArrayProvided();
  console.log('prepared billing cache tests passed');
}

run();
