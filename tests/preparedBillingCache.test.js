const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const mainCode = fs.readFileSync(path.join(__dirname, '../src/main.gs'), 'utf8');

function buildValidPreparedPayload(ctx, overrides = {}) {
  return Object.assign({
    schemaVersion: ctx.PREPARED_BILLING_SCHEMA_VERSION || 2,
    billingMonth: '202501',
    billingJson: [{ patientId: 'P001', billingMonth: '202501' }],
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

function createBankSheetMock(name, rows = []) {
  const data = rows.map(row => row.slice());
  const sheet = {
    _name: name,
    _index: 1,
    _workbook: null,
    _rows: data,
    getName: () => sheet._name,
    setName: newName => {
      const oldName = sheet._name;
      sheet._name = newName;
      if (sheet._workbook && typeof sheet._workbook._renameSheet === 'function') {
        sheet._workbook._renameSheet(oldName, sheet);
      }
    },
    getIndex: () => sheet._index,
    setIndex: newIndex => { sheet._index = newIndex; },
    getLastRow: () => sheet._rows.length,
    getLastColumn: () => sheet._rows.reduce((max, row) => Math.max(max, row.length), 0),
    getRange: (row, col, numRows = 1, numCols = 1) => {
      const getValues = () => {
        const values = [];
        for (let r = 0; r < numRows; r++) {
          const sourceRow = sheet._rows[row - 1 + r] || [];
          const rowValues = [];
          for (let c = 0; c < numCols; c++) {
            rowValues.push(sourceRow[col - 1 + c] !== undefined ? sourceRow[col - 1 + c] : '');
          }
          values.push(rowValues);
        }
        return values;
      };

      const applyValues = values => {
        for (let r = 0; r < numRows; r++) {
          const destIndex = row - 1 + r;
          sheet._rows[destIndex] = sheet._rows[destIndex] || [];
          for (let c = 0; c < numCols; c++) {
            sheet._rows[destIndex][col - 1 + c] = values[r][c];
          }
        }
      };

      return {
        getValues,
        getDisplayValues: () => getValues(),
        setValues: applyValues,
        setValue: value => applyValues([[value]])
      };
    },
    insertColumnAfter: colIndex => {
      const insertAt = Math.max(0, colIndex);
      sheet._rows = sheet._rows.map(row => {
        const copy = row.slice();
        copy.splice(insertAt, 0, '');
        return copy;
      });
    },
    getProtections: () => [],
    copyTo: workbook => {
      const copy = createBankSheetMock(name, sheet._rows);
      workbook._registerSheet(copy);
      return copy;
    }
  };
  return sheet;
}

function createBankWorkbookMock(templateSheet) {
  const sheets = {};
  const order = [];
  const workbook = {};
  const register = sheet => {
    const index = order.length + 1;
    sheet.setIndex(index);
    sheet._workbook = workbook;
    order.push(sheet);
    sheets[sheet.getName()] = sheet;
    return sheet;
  };

  Object.assign(workbook, {
    _registerSheet: register,
    _renameSheet: (oldName, sheet) => {
      if (oldName && sheets[oldName] === sheet) {
        delete sheets[oldName];
      }
      sheets[sheet.getName()] = sheet;
    },
    getSheetByName: name => sheets[name] || null,
    deleteSheet: sheet => {
      const name = sheet.getName();
      delete sheets[name];
      const idx = order.indexOf(sheet);
      if (idx >= 0) order.splice(idx, 1);
    },
    getNumSheets: () => order.length,
    setActiveSheet: () => {},
    moveActiveSheet: () => {},
    insertSheet: name => register(createBankSheetMock(name)),
    getActiveSheet: () => null
  });

  if (templateSheet) {
    register(templateSheet);
  }

  return workbook;
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
  const ctx = createMainContext({
    normalizeBillingMonthInput: input => ({ key: String(input) }),
    loadPreparedBillingFromSheet_: () => buildValidPreparedPayload(ctx, { billingJson: [], billingMonth: '202501' }),
    normalizePreparedBilling_: payload => payload,
    exportBankTransferDataForPrepared_: () => assert.fail('invalid payload should not reach exporter')
  });

  assert.throws(() => {
    ctx.generateBankTransferDataFromCache('202501');
  }, /破損しています/);
}

function testBankExportReportsLedgerIssues() {
  let ctx;
  ctx = createMainContext({
    normalizeBillingMonthInput: input => ({ key: String(input) }),
    loadPreparedBillingFromSheet_: () => ({ billingJson: [{ patientId: 'P001', billingMonth: '202501' }], billingMonth: '202501', schemaVersion: 2,
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

function testPreparedBillingSheetFallbackRestoresCache() {
  let savedPayload = null;
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
    savePreparedBilling_: payload => { savedPayload = payload; },
    billingLogger_: { log: () => {} }
  });

  ctx.loadPreparedBillingWithSheetFallback_('2025-01', { withValidation: true });

  assert.ok(savedPayload, 'fallback payload is restored to cache');
  assert.strictEqual(savedPayload.billingMonth, '202501');
}

function testPreparedBillingCacheIsChunkedWhenTooLarge() {
  const store = {};
  const cache = {
    get: key => store[key] || null,
    put: (key, value) => { store[key] = value; },
    remove: key => { delete store[key]; },
    removeAll: keys => { (keys || []).forEach(k => delete store[k]); }
  };

  const ctx = createMainContext({
    CacheService: { getScriptCache: () => cache },
    billingLogger_: { log: () => {} },
    normalizeBillingMonthKeySafe_: value => String(value || '').replace(/\D/g, ''),
    normalizePreparedBilling_: payload => payload
  });

  const chunkLimit = ctx.BILLING_CACHE_MAX_ENTRY_LENGTH || 90000;
  const chunkMarkerPrefix = ctx.BILLING_CACHE_CHUNK_MARKER || 'chunked:';
  const largeText = 'X'.repeat(chunkLimit + 5000);
  const payload = buildValidPreparedPayload(ctx, {
    billingMonth: '202501',
    billingJson: [{ patientId: 'P001', memo: largeText }]
  });

  ctx.savePreparedBilling_(payload);

  const cacheKey = ctx.buildBillingCacheKey_('202501');
  const chunkMarker = store[cacheKey];
  const chunkCount = parseInt(String(chunkMarker).slice(chunkMarkerPrefix.length), 10);

  assert.ok(String(chunkMarker || '').startsWith(chunkMarkerPrefix), 'cache uses chunk marker when oversized');
  assert.ok(chunkCount > 1, 'payload is split into multiple chunks');
  assert.ok(store[ctx.buildBillingCacheChunkKey_(cacheKey, 1)], 'first chunk is stored');

  const loaded = ctx.loadPreparedBilling_('202501');

  assert.ok(loaded && Array.isArray(loaded.billingJson), 'chunked cache is reconstructed');
  assert.strictEqual(loaded.billingJson[0].memo.length, largeText.length, 'chunked fields are preserved');
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
  const metaJsonSheet = sheets.PreparedBillingMetaJson;
  const jsonSheet = sheets.PreparedBillingJson;
  assert.ok(metaSheet && jsonSheet && metaJsonSheet, 'Meta/Json シートが作成される');
  assert.deepStrictEqual(metaSheet._rows[0], ['billingMonth', 'preparedAt', 'preparedBy', 'payloadVersion', 'note'], 'メタシートのヘッダーが作成される');
  assert.deepStrictEqual(metaJsonSheet._rows[0], ['billingMonth', 'chunkIndex', 'payloadChunk'], 'メタJSONシートのヘッダーが作成される');
  assert.deepStrictEqual(jsonSheet._rows[0], ['billingMonth', 'patientId', 'billingRowJson'], '明細シートのヘッダーが作成される');
  assert.strictEqual(metaSheet._rows[1][0], '202501', 'billingMonthがメタに保存される');
  assert.ok(String(metaSheet._rows[1][2]).includes('user@example.com'), '実行ユーザーが保存される');
  assert.strictEqual(metaJsonSheet._rows.length > 1, true, 'メタJSONが分割保存される');
  assert.strictEqual(jsonSheet._rows.length, 3, 'billingJsonの件数分保存される');

  const loaded = ctx.loadPreparedBillingFromSheet_('202501');
  assert.ok(loaded, '保存したデータを読み込める');
  assert.strictEqual(loaded.billingMonth, '202501');
  assert.strictEqual(loaded.preparedAt, '2025-01-02T03:04:05.000Z');
  assert.strictEqual(loaded.schemaVersion, 2);
  assert.strictEqual(Array.isArray(loaded.billingJson) ? loaded.billingJson.length : 0, 2, 'billingJsonが復元される');
}

function testBankWithdrawalSheetRegeneration() {
  const bankTemplate = createBankSheetMock('銀行情報', [
    ['名前', '金額', '口座'],
    ['田中 太郎', '', '1234']
  ]);
  const workbook = createBankWorkbookMock(bankTemplate);
  const columnLetterToNumber_ = letter => {
    const normalized = String(letter || '').toUpperCase();
    let result = 0;
    for (let i = 0; i < normalized.length; i++) {
      const code = normalized.charCodeAt(i);
      if (code < 65 || code > 90) continue;
      result = result * 26 + (code - 64);
    }
    return result;
  };

  const ctx = createMainContext({
    BILLING_LABELS: { name: ['名前'], furigana: ['フリガナ'] },
    billingLogger_: { log: () => {} },
    normalizeBillingMonthInput: () => ({ key: '202404', year: 2024, month: 4 }),
    billingNormalizePatientId_: value => String(value || '').trim(),
    normalizeBillingNameKey_: value => String(value || '').replace(/\s+/g, '').trim(),
    normalizeMoneyNumber_: value => Number(value) || 0,
    resolveBillingColumn_: (headers, candidates, _label, options) => {
      const normalized = headers.map(header => String(header || '').trim());
      const candidateList = Array.isArray(candidates) ? candidates : [];
      for (let i = 0; i < candidateList.length; i++) {
        const idx = normalized.indexOf(candidates[i]);
        if (idx >= 0) return idx + 1;
      }
      if (options && options.fallbackLetter) return columnLetterToNumber_(options.fallbackLetter);
      if (options && options.fallbackIndex) return options.fallbackIndex;
      return 0;
    },
    SpreadsheetApp: {
      ProtectionType: { SHEET: 'SHEET' },
      getActiveSpreadsheet: () => workbook
    },
    columnLetterToNumber_
  });
  ctx.ensureBankInfoSheet_ = () => bankTemplate;
  ctx.billingSs = () => workbook;
  ctx.ss = () => workbook;

  const prepared = {
    billingMonth: '202404',
    patients: {
      P001: { nameKanji: '田中 太郎' },
      P002: { nameKanji: '新規 花子' }
    },
    billingJson: [
      { patientId: 'P001', grandTotal: 1000 },
      { patientId: 'P002', grandTotal: 2000 }
    ]
  };

  ctx.syncBankWithdrawalSheetForMonth_('202404', prepared);

  const firstSheet = workbook.getSheetByName('銀行引落_2024-04');
  assert.ok(firstSheet, '初回生成で月次シートが作成される');
  const firstValues = firstSheet.getRange(1, 1, firstSheet.getLastRow(), firstSheet.getLastColumn()).getValues();
  assert.strictEqual(firstValues.length, 2, 'テンプレート行数を引き継ぐ');
  const firstAmountIndex = firstValues[0].indexOf('金額');
  assert.ok(firstAmountIndex >= 0, '金額列が存在する');
  assert.strictEqual(firstValues[1][firstAmountIndex], 1000, '請求金額が金額列に出力される');

  // Add a new account to the template and alter the existing monthly sheet to ensure it is refreshed.
  bankTemplate.getRange(3, 1, 1, 3).setValues([['新規 花子', '', '5678']]);
  firstSheet.getRange(2, 2, 1, 1).setValues([[9999]]);

  ctx.syncBankWithdrawalSheetForMonth_('202404', prepared);

  const refreshedSheet = workbook.getSheetByName('銀行引落_2024-04');
  const refreshedValues = refreshedSheet.getRange(1, 1, refreshedSheet.getLastRow(), refreshedSheet.getLastColumn()).getValues();
  assert.strictEqual(refreshedValues.length, 3, 'テンプレートの新規行が再生成で反映される');
  const refreshedAmountIndex = refreshedValues[0].indexOf('金額');
  assert.ok(refreshedAmountIndex >= 0, '再生成後も金額列が存在する');
  assert.strictEqual(refreshedValues[1][refreshedAmountIndex], 1000, '既存利用者の金額は再生成後も正しく計算される');
  assert.strictEqual(refreshedValues[2][refreshedAmountIndex], 2000, '再生成で新規利用者の金額も反映される');

  const templateValues = bankTemplate.getRange(1, 1, bankTemplate.getLastRow(), bankTemplate.getLastColumn()).getValues();
  assert.strictEqual(templateValues[1][1], '', '銀行情報シートは参照専用で金額が書き換わらない');
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
  testPreparedBillingSheetFallbackRestoresCache();
  testPreparedBillingCacheIsChunkedWhenTooLarge();
  testInvoiceGenerationUsesSheetWhenCacheMissing();
  testPrepareBillingDataNormalizesMonthKey();
  testPreparedBillingSheetSaveAndLoad();
  testBankWithdrawalSheetRegeneration();
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
    ensureBankWithdrawalSheet_: () => ({ getLastRow: () => 0 }),
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
