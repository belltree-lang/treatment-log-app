const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');

const treatmentRows = [
  { ts: new Date(2025, 1, 4, 10, 0, 0), pid: 'P001', note: 'feb', email: '', treatmentId: 'FEB1', category: '', monthKey: '202502' },
  { ts: new Date(2025, 0, 28, 10, 0, 0), pid: 'P001', note: 'jan', email: '', treatmentId: 'JAN1', category: '', monthKey: '202501' },
  { ts: new Date(2025, 0, 10, 10, 0, 0), pid: 'P002', note: 'other', email: '', treatmentId: 'OTH1', category: '', monthKey: '202501' }
];

let col13ReadCount = 0;
let rowScopedColumnReads = 0;

const sheet = {
  getLastRow: () => treatmentRows.length + 1,
  getRange: (row, col, numRows) => {
    const values = [];
    for (let i = 0; i < numRows; i++) {
      const record = treatmentRows[(row - 2) + i];
      let value = '';
      switch (col) {
        case 1: value = record.ts; break;
        case 2: value = record.pid; break;
        case 3: value = record.note; break;
        case 4: value = record.email; break;
        case 7: value = record.treatmentId; break;
        case 8: value = record.category; break;
        case 13: value = record.monthKey; break;
        default: value = '';
      }
      values.push([value]);
    }
    if (col === 13) col13ReadCount += 1;
    if ([1, 2, 3, 4, 7, 8].includes(col)) {
      rowScopedColumnReads += numRows;
    }
    return {
      getValues: () => values,
      getDisplayValues: () => values.map(rowVals => rowVals.map(v => (v == null ? '' : String(v)))),
      setValues: nextValues => {
        if (col === 7) {
          for (let i = 0; i < nextValues.length; i++) {
            treatmentRows[(row - 2) + i].treatmentId = nextValues[i][0];
          }
        }
      }
    };
  }
};

function formatDate(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const HH = String(date.getHours()).padStart(2, '0');
  const MM = String(date.getMinutes()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd} ${HH}:${MM}`;
}

const context = {
  normId_: v => (v ? String(v).trim() : ''),
  resolveYearMonthOrCurrent_: (year, month) => ({ year: Number(year), month: Number(month) }),
  PATIENT_CACHE_KEYS: { treatments: pid => pid },
  PATIENT_CACHE_TTL_SECONDS: 0,
  sh: () => sheet,
  Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
  Utilities: { formatDate },
  mapTreatmentCategoryCellToKey_: label => (label ? String(label) : ''),
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
  Logger: { log: () => {} }
};

vm.createContext(context);
vm.runInContext(code, context);
let capturedKeys = [];
context.cacheFetch_ = (key, fn) => {
  capturedKeys.push(key);
  return fn();
};

const janRows = JSON.parse(JSON.stringify(context.listTreatmentsForMonth('P001', 2025, 1)));
const febRows = JSON.parse(JSON.stringify(context.listTreatmentsForMonth('P001', 2025, 2)));

assert.deepStrictEqual(janRows.map(r => r.treatmentId), ['JAN1'], '指定月のデータだけを返す');
assert.deepStrictEqual(febRows.map(r => r.treatmentId), ['FEB1'], '別月のデータも取得できる');
assert.ok(capturedKeys.includes('patient:treatments:P001:202501'), 'キャッシュキーにYYYYMMを含める');
assert.ok(capturedKeys.includes('patient:treatments:P001:202502'), '月ごとにキャッシュキーを分離する');
assert.strictEqual(col13ReadCount, 2, '月検索ではM列を都度読み取る');
assert.strictEqual(rowScopedColumnReads, 18, '一致行のみ必要列を読み取る');

console.log('treatment list month filter tests passed');
