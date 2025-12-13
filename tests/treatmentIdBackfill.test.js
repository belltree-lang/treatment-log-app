const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');

const now = new Date();
const makeDate = day => new Date(now.getFullYear(), now.getMonth(), day, 12, 0, 0);

const treatmentRows = [
  { ts: makeDate(3), pid: 'P001', note: 'with id', email: 'with@example.com', treatmentId: 'T-1', category: '30分施術（保険）' },
  { ts: makeDate(4), pid: 'P001', note: 'missing id', email: 'missing@example.com', treatmentId: '', category: '30分施術（保険）' },
];

const columnData = {
  1: treatmentRows.map(r => r.ts),
  2: treatmentRows.map(r => r.pid),
  3: treatmentRows.map(r => r.note),
  4: treatmentRows.map(r => r.email),
  7: treatmentRows.map(r => r.treatmentId),
  8: treatmentRows.map(r => r.category),
};

let lastWrittenTreatmentIds = null;

const sheet = {
  getLastRow: () => treatmentRows.length + 1,
  getRange: (row, col, numRows) => {
    const source = columnData[col] || [];
    const slice = [];
    for (let i = 0; i < numRows; i++) {
      slice.push([source[i]]);
    }
    return {
      getValues: () => slice.map(r => r.slice()),
      getDisplayValues: () => slice.map(r => r.map(val => (val == null ? '' : String(val)))),
      setValues: values => {
        lastWrittenTreatmentIds = values.map(rowVals => rowVals[0]);
        columnData[col] = lastWrittenTreatmentIds.slice();
      },
    };
  }
};

const context = {
  normId_: v => (v ? String(v).trim() : ''),
  PATIENT_CACHE_KEYS: { treatments: pid => pid },
  PATIENT_CACHE_TTL_SECONDS: 0,
  sh: () => sheet,
  Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
  Utilities: {
    formatDate: date => date.toISOString().slice(0, 16).replace('T', ' '),
    getUuid: () => 'generated-uuid'
  },
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

assert.strictEqual(normalizedRows.length, 2, 'returns all rows for the patient');
assert.deepStrictEqual(
  normalizedRows.map(r => r.treatmentId).sort(),
  ['T-1', 'generated-uuid'].sort(),
  'fills missing treatmentId values'
);
assert.deepStrictEqual(
  lastWrittenTreatmentIds,
  ['T-1', 'generated-uuid'],
  'writes back generated treatmentId to the sheet'
);

console.log('treatmentId backfill tests passed');
