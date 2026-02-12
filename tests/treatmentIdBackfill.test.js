const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');

const now = new Date();
const makeDate = day => new Date(now.getFullYear(), now.getMonth(), day, 12, 0, 0);
const monthKey = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}`;

const treatmentRows = [
  { ts: makeDate(3), pid: 'P001', note: 'with id', email: 'with@example.com', treatmentId: 'T-1', category: '30分施術（保険）', monthKey },
  { ts: makeDate(4), pid: 'P001', note: 'missing id', email: 'missing@example.com', treatmentId: '', category: '30分施術（保険）', monthKey },
];

let lastWrittenTreatmentIds = null;

const sheet = {
  getLastRow: () => treatmentRows.length + 1,
  getRange: (row, col, numRows) => {
    const slice = [];
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
      slice.push([value]);
    }
    return {
      getValues: () => slice.map(r => r.slice()),
      getDisplayValues: () => slice.map(r => r.map(val => (val == null ? '' : String(val)))),
      setValues: values => {
        if (col !== 7) return;
        lastWrittenTreatmentIds = values.map(rowVals => rowVals[0]);
        for (let i = 0; i < values.length; i++) {
          treatmentRows[(row - 2) + i].treatmentId = values[i][0];
        }
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
