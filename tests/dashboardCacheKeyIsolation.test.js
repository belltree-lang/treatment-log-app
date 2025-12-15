const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const sheetUtilsCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'utils', 'sheetUtils.js'), 'utf8');
const cacheUtilsCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'utils', 'cacheUtils.js'), 'utf8');
const configCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'config.gs'), 'utf8');
const loadInvoicesCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'data', 'loadInvoices.js'), 'utf8');
const loadTreatmentLogsCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'data', 'loadTreatmentLogs.js'), 'utf8');

function createContext(overrides = {}) {
  const context = {
    console,
    JSON,
    Date,
    Set,
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        if (fmt === 'yyyyMM') return date.toISOString().slice(0, 7).replace('-', '');
        if (fmt === 'yyyy年MM月') {
          const iso = date.toISOString();
          return `${iso.slice(0, 4)}年${iso.slice(5, 7)}月`;
        }
        return date.toISOString();
      }
    },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' }
  };
  vm.createContext(context);
  vm.runInContext(configCode, context);
  vm.runInContext(sheetUtilsCode, context);
  vm.runInContext(cacheUtilsCode, context);
  vm.runInContext(loadInvoicesCode, context);
  vm.runInContext(loadTreatmentLogsCode, context);
  Object.assign(context, overrides);
  return context;
}

function testInvoiceCacheKeyIncludesMonth() {
  const keys = [];
  const ctx = createContext({
    dashboardCacheFetch_: (key, fetchFn) => { keys.push(key); return fetchFn(); },
    dashboardGetInvoiceRootFolder_: () => null
  });

  ctx.loadInvoices({ patientInfo: { patients: {}, warnings: [] }, now: new Date('2025-03-01T00:00:00Z') });
  ctx.loadInvoices({ patientInfo: { patients: {}, warnings: [] }, now: new Date('2025-04-01T00:00:00Z') });

  assert.notStrictEqual(keys[0], keys[1], '月が異なる場合はキャッシュキーも変わる');
  assert.ok(keys[0].includes('202503'), 'キャッシュキーが対象月を含む');
  assert.ok(keys[1].includes('202504'), 'キャッシュキーが対象月を含む');
}

function testTreatmentLogsCacheKeyIncludesMonth() {
  const keys = [];
  const ctx = createContext({
    dashboardCacheFetch_: (key, fetchFn) => { keys.push(key); return fetchFn(); },
    dashboardGetSpreadsheet_: () => null
  });

  ctx.loadTreatmentLogs({ patientInfo: { patients: {}, nameToId: {}, warnings: [] }, now: new Date('2025-02-15T00:00:00Z') });
  ctx.loadTreatmentLogs({ patientInfo: { patients: {}, nameToId: {}, warnings: [] }, now: new Date('2025-03-01T00:00:00Z') });

  assert.notStrictEqual(keys[0], keys[1], '施術録も月が変わればキャッシュキーが分離される');
  assert.ok(keys[0].includes('202502'), '施術録キャッシュキーが対象月を含む');
  assert.ok(keys[1].includes('202503'), '施術録キャッシュキーが対象月を含む');
}

(function run() {
  testInvoiceCacheKeyIncludesMonth();
  testTreatmentLogsCacheKeyIncludesMonth();
  console.log('dashboard cache key isolation tests passed');
})();
