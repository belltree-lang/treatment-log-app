const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');

const now = new Date();
const makeDate = day => new Date(now.getFullYear(), now.getMonth(), day, 12, 0, 0);
const monthKey = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}`;

const treatmentRows = [
  { ts: makeDate(5), pid: 'P001', note: 'newest', email: 'latest@example.com', treatmentId: 'T2', category: '30分施術（保険）', monthKey },
  { ts: makeDate(2), pid: 'P001', note: 'older', email: 'older@example.com', treatmentId: 'T1', category: '60分施術（混合）', monthKey },
  { ts: makeDate(1), pid: 'P002', note: 'other patient', email: 'other@example.com', treatmentId: 'X1', category: '30分施術（保険）', monthKey }
];

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

    return {
      getValues: () => values,
      getDisplayValues: () => values.map(rowVals => rowVals.map(val => (val == null ? '' : String(val)))),
      setValues: () => {}
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
  PATIENT_CACHE_KEYS: { treatments: pid => pid },
  PATIENT_CACHE_TTL_SECONDS: 0,
  sh: () => sheet,
  Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
  Utilities: { formatDate },
  mapTreatmentCategoryCellToKey_: label => (label ? String(label) : ''),
  CacheService: {
    getScriptCache: () => ({ get: () => null, put: () => {} })
  },
  Logger: { log: () => {} }
};

vm.createContext(context);
vm.runInContext(code, context);

context.cacheFetch_ = (key, fn) => fn();

const rows = context.listTreatmentsForCurrentMonth('P001');
const normalizedRows = JSON.parse(JSON.stringify(rows));

assert.strictEqual(normalizedRows.length, 2, 'filters to the target patient');
assert.deepStrictEqual(normalizedRows.map(r => r.treatmentId), ['T2', 'T1'], 'sorted by timestamp in descending order');
assert.ok(/T2/.test(normalizedRows[0].treatmentId) && normalizedRows[0].when.includes('-'), 'keeps formatted date text');

console.log('treatment list ordering tests passed');
