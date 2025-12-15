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
  const baseSpreadsheet = {
    getSheetByName: () => null,
    insertSheet: () => ({
      getRange: () => ({ setValues: () => {} }),
      getLastRow: () => 0
    })
  };
  const ctx = Object.assign({
    console,
    CacheService: baseCache,
    Session: {},
    SpreadsheetApp: { getActiveSpreadsheet: () => baseSpreadsheet }
  }, overrides);
  vm.createContext(ctx);
  vm.runInContext(mainCode, ctx);
  return Object.assign(ctx, overrides);
}

function createSheetMock() {
  const rows = [];
  return {
    _rows: rows,
    getLastRow: () => rows.length,
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => {
        const values = [];
        for (let r = 0; r < numRows; r++) {
          const targetRow = rows[row - 1 + r] || [];
          const rowValues = [];
          for (let c = 0; c < numCols; c++) {
            rowValues.push(targetRow[col - 1 + c] !== undefined ? targetRow[col - 1 + c] : '');
          }
          values.push(rowValues);
        }
        return values;
      },
      setValues: values => {
        for (let r = 0; r < numRows; r++) {
          const destIndex = row - 1 + r;
          rows[destIndex] = rows[destIndex] || [];
          for (let c = 0; c < numCols; c++) {
            rows[destIndex][col - 1 + c] = values[r][c];
          }
        }
      },
      clearContent: () => {
        for (let r = 0; r < numRows; r++) {
          const destIndex = row - 1 + r;
          if (rows[destIndex]) {
            for (let c = 0; c < numCols; c++) {
              rows[destIndex][col - 1 + c] = '';
            }
          }
        }
      }
    }),
    deleteRow: rowIndex => {
      rows.splice(rowIndex - 1, 1);
    },
    insertRows: (rowIndex, howMany) => {
      const inserts = new Array(howMany).fill(null).map(() => []);
      rows.splice(rowIndex - 1, 0, ...inserts);
    }
  };
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
    loadPreparedBillingFromSheet_: () => ({
      billingMonth: '202501',
      billingJson: {},
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
    }),
    normalizePreparedBilling_: payload => payload,
    exportBankTransferDataForPrepared_: () => assert.fail('invalid payload should not reach exporter')
  });

  assert.throws(() => {
    ctx.generateBankTransferDataFromCache('202501');
  }, /検証に失敗/);
}

function testBankExportPassesWhenArrayProvided() {
  const exportCalls = [];
  let ctx;
  ctx = createMainContext({
    normalizeBillingMonthInput: input => ({ key: String(input) }),
    loadPreparedBillingFromSheet_: () => buildValidPreparedPayload(ctx, {
      billingJson: [{ billingMonth: '202501', patientId: 'p1' }],
      billingMonth: '202501'
    }),
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

function testBankExportAcceptsYmObject() {
  const exportCalls = [];
  let ctx;
  ctx = createMainContext({
    normalizeBillingMonthInput: input => {
      const raw = typeof input === 'object' && input && input.ym ? input.ym : input;
      return { key: String(raw).replace(/\D/g, '') };
    },
    normalizePreparedBilling_: payload => payload,
    loadPreparedBillingFromSheet_: monthKey => {
      exportCalls.push(monthKey);
      return buildValidPreparedPayload(ctx, {
        billingJson: [{ billingMonth: '202501', patientId: 'p1' }],
        billingMonth: '202501'
      });
    },
    exportBankTransferDataForPrepared_: prepared => ({
      billingMonth: prepared.billingMonth,
      inserted: prepared.billingJson.length
    })
  });

  const result = ctx.generateBankTransferDataFromCache({ ym: '2025-01' }, { billingMonth: { ym: '2025-01' } });

  assert.deepStrictEqual(exportCalls, ['202501']);
  assert.strictEqual(result.billingMonth, '202501');
  assert.strictEqual(result.inserted, 1);
}

function testBankExportReturnsEmptyForZeroBilling() {
  const exportCalls = [];
  let ctx;
  ctx = createMainContext({
    normalizeBillingMonthInput: input => ({ key: String(input) }),
    loadPreparedBillingFromSheet_: () => buildValidPreparedPayload(ctx, { billingJson: [], billingMonth: '202501' }),
    normalizePreparedBilling_: payload => payload,
    exportBankTransferDataForPrepared_: () => exportCalls.push('called')
  });

  const result = ctx.generateBankTransferDataFromCache('202501');
  assert.deepStrictEqual(exportCalls, [], '0件のときはエクスポート処理を呼ばない');
  assert.strictEqual(result.billingMonth, '202501');
  assert.ok(Array.isArray(result.rows) && result.rows.length === 0, 'rows should be empty array');
  assert.strictEqual(result.inserted, 0);
  assert.strictEqual(result.skipped, 0);
  assert.strictEqual(result.message, '当月の請求対象はありません', '0件の場合は案内メッセージを返す');
}

function testBankExportReportsLedgerIssues() {
  let ctx;
  ctx = createMainContext({
    normalizeBillingMonthInput: input => ({ key: String(input) }),
    loadPreparedBillingFromSheet_: () => ({ billingJson: [], billingMonth: '202501', schemaVersion: 2,
      carryOverLedgerByPatient: {}, carryOverLedgerMeta: {}, visitsByPatient: {}, totalsByPatient: {}, patients: {},
      bankInfoByName: {}, staffByPatient: {}, staffDirectory: {}, staffDisplayByPatient: {}, billingOverrideFlags: {},
      carryOverByPatient: {}, bankAccountInfoByPatient: {} }),
    normalizePreparedBilling_: payload => payload
  });

  assert.throws(() => {
    ctx.generateBankTransferDataFromCache('202501');
  }, /繰越金データを確認してください/);
}

function testBankExportReportsMissingPreparation() {
  const ctx = createMainContext({
    normalizeBillingMonthInput: input => ({ key: String(input) }),
    loadPreparedBillingFromSheet_: () => null,
    normalizePreparedBilling_: () => null
  });

  assert.throws(() => {
    ctx.generateBankTransferDataFromCache('202501');
  }, /集計・確定/);
}

function testPreparedBillingSheetFallback() {
  let cacheCalls = 0;
  const ctx = createMainContext({
    normalizeBillingMonthInput: input => ({ key: String(input).replace(/\D/g, '') || '' }),
    loadPreparedBilling_: () => { cacheCalls += 1; return null; },
    loadPreparedBillingFromSheet_: () => buildValidPreparedPayload(ctx, {
      billingMonth: '202501',
      billingJson: [{ billingMonth: '202501', patientId: 'p1' }]
    }),
    normalizePreparedBilling_: payload => payload,
    validatePreparedBillingPayload_: payload => payload
      ? { ok: true, billingMonth: payload.billingMonth }
      : { ok: false, reason: 'missing' }
  });

  const result = ctx.loadPreparedBillingWithSheetFallback_('2025-01', { withValidation: true });
  assert.strictEqual(cacheCalls, 1, 'cache lookup is attempted first');
  assert.ok(result && result.prepared, 'fallback returns prepared payload');
  assert.strictEqual(result.prepared.billingMonth, '202501', 'sheet payload billingMonth is preserved');
  assert.ok(result.validation && result.validation.ok, 'validation result is returned');
}

function testInvoiceGenerationUsesSheetWhenCacheMissing() {
  const invoiceCalls = [];
  const ctx = createMainContext({
    normalizeBillingMonthInput: input => ({ key: String(input).replace(/\D/g, '') || '' }),
    loadPreparedBilling_: () => null,
    loadPreparedBillingFromSheet_: () => buildValidPreparedPayload(ctx, {
      billingMonth: '202501',
      billingJson: [{ billingMonth: '202501', patientId: 'p1' }]
    }),
    normalizePreparedBilling_: payload => payload,
    validatePreparedBillingPayload_: payload => payload
      ? { ok: true, billingMonth: payload.billingMonth }
      : { ok: false, reason: 'missing' },
    generatePreparedInvoices_: prepared => {
      invoiceCalls.push(prepared.billingMonth);
      return { billingMonth: prepared.billingMonth, billingJson: prepared.billingJson };
    }
  });

  const result = ctx.generateInvoicesFromCache('2025/01');
  assert.deepStrictEqual(invoiceCalls, ['202501'], 'invoice generation uses normalized month from sheet payload');
  assert.strictEqual(result && result.billingMonth, '202501');
}

function testPreparedBillingSheetSaveAndLoad() {
  const sheets = {};
  const spreadsheet = {
    getSheetByName: name => sheets[name] || null,
    insertSheet: name => {
      sheets[name] = createSheetMock();
      return sheets[name];
    }
  };
  const ctx = createMainContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Session: { getActiveUser: () => ({ getEmail: () => 'user@example.com' }) },
    billingLogger_: { log: () => {} }
  });

  const payload = buildValidPreparedPayload(ctx, {
    preparedAt: '2025-01-02T03:04:05.000Z',
    billingJson: [
      { patientId: 'P001', billingMonth: '202501' },
      { patientId: 'P002', billingMonth: '202501' }
    ]
  });

  const saveResult = ctx.savePreparedBillingToSheet_('202501', payload);
  assert.strictEqual(saveResult && saveResult.billingMonth, '202501');
  const metaSheet = sheets.PreparedBillingMeta;
  const jsonSheet = sheets.PreparedBillingJson;
  assert.ok(metaSheet && jsonSheet, 'Meta/Json シートが作成される');
  assert.deepStrictEqual(metaSheet._rows[0], ['billingMonth', 'preparedAt', 'preparedBy', 'payloadVersion', 'payloadJson', 'note'], 'メタシートのヘッダーが作成される');
  assert.deepStrictEqual(jsonSheet._rows[0], ['billingMonth', 'patientId', 'billingRowJson'], '明細シートのヘッダーが作成される');
  assert.strictEqual(metaSheet._rows[1][0], '202501', 'billingMonthがメタに保存される');
  assert.ok(String(metaSheet._rows[1][2]).includes('user@example.com'), '実行ユーザーが保存される');
  assert.strictEqual(jsonSheet._rows.length, 3, 'billingJsonの件数分保存される');

  const loaded = ctx.loadPreparedBillingFromSheet_('202501');
  assert.ok(loaded, '保存したデータを読み込める');
  assert.strictEqual(loaded.billingMonth, '202501');
  assert.strictEqual(loaded.preparedAt, '2025-01-02T03:04:05.000Z');
  assert.strictEqual(loaded.schemaVersion, 2);
  assert.strictEqual(Array.isArray(loaded.billingJson) ? loaded.billingJson.length : 0, 2, 'billingJsonが復元される');
}

function run() {
  testValidationRequiresLedgerFields();
  testValidationAcceptsCompletePayload();
  testSchemaVersionIsValidated();
  testBankExportRejectsNonArrayBillingJson();
  testBankExportPassesWhenArrayProvided();
  testBankExportAcceptsYmObject();
  testBankExportReturnsEmptyForZeroBilling();
  testBankExportReportsLedgerIssues();
  testBankExportReportsMissingPreparation();
  testPreparedBillingSheetFallback();
  testInvoiceGenerationUsesSheetWhenCacheMissing();
  testPrepareBillingDataNormalizesMonthKey();
  testPreparedBillingSheetSaveAndLoad();
  console.log('prepared billing cache tests passed');
}

run();

function testPrepareBillingDataNormalizesMonthKey() {
  const savedPayloads = [];
  const savedSheetPayloads = [];
  const ctx = createMainContext({
    normalizeBillingMonthInput: input => ({
      key: String(input).replace(/\D/g, ''),
      start: new Date(),
      end: new Date(),
      timezone: 'Asia/Tokyo',
      year: 2025,
      month: 1
    }),
    billingLogger_: { log: () => {} },
    buildPreparedBillingPayload_: month => ({
      billingMonth: month,
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
    }),
    toClientBillingPayload_: payload => payload,
    serializeBillingPayload_: payload => payload,
    savePreparedBilling_: payload => savedPayloads.push(payload),
    savePreparedBillingToSheet_: (billingMonth, payload) => savedSheetPayloads.push(Object.assign({}, payload, { billingMonth }))
  });

  const result = ctx.prepareBillingData('2025-01');

  assert.strictEqual(result.billingMonth, '202501', '返却されるbillingMonthはYYYYMMに正規化される');
  assert.strictEqual(savedPayloads.length, 1, 'prepared payloadが保存される');
  assert.strictEqual(savedPayloads[0].billingMonth, '202501', '保存されるpayloadのbillingMonthも正規化される');
  assert.strictEqual(savedSheetPayloads.length, 1, 'シート保存用のpayloadも保存される');
  assert.strictEqual(savedSheetPayloads[0].billingMonth, '202501', 'シート保存payloadのbillingMonthも正規化される');
}
