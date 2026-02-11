const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');
const sandbox = { console };

function formatYearMonth(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  return `${yyyy}${mm}`;
}
vm.createContext(sandbox);
vm.runInContext(code, sandbox);

function createSheet(rows) {
  return {
    rows: rows.map(r => r.slice()),
    getLastRow() { return this.rows.length; },
    getMaxColumns() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); },
    getRange(row, col, numRows = 1, numCols = 1) {
      const self = this;
      return {
        getValues() {
          const out = [];
          for (let i = 0; i < numRows; i++) {
            const r = self.rows[row - 1 + i] || [];
            const slice = [];
            for (let j = 0; j < numCols; j++) {
              slice.push(r[col - 1 + j]);
            }
            out.push(slice);
          }
          return out;
        },
        getValue() {
          const r = self.rows[row - 1] || [];
          return r[col - 1];
        },
        setValue(val) {
          const idx = row - 1;
          if (!self.rows[idx]) self.rows[idx] = [];
          self.rows[idx][col - 1] = val;
        }
      };
    },
    deleteRow(rowNumber) {
      this.rows.splice(rowNumber - 1, 1);
    }
  };
}

function testDeleteUsesTreatmentId() {
  const now = new Date();
  const d1 = new Date(now.getFullYear(), now.getMonth(), 1, 9, 0, 0);
  const d2 = new Date(now.getFullYear(), now.getMonth(), 2, 9, 0, 0);
  const d3 = new Date(now.getFullYear(), now.getMonth(), 3, 9, 0, 0);
  const sheet = createSheet([
    ['TS', '患者ID', '所見', 'email', '', '', '施術ID'],
    [d1, '0001', 'first', 'a@example.com', '', '', 'tid-a'],
    [d2, '0001', 'second', '', '', '', 'tid-b'],
    [d3, '0001', 'third', '', '', '', 'tid-c']
  ]);

  const cleared = [];
  const invalidated = [];
  const fetched = [];

  sandbox.sh = () => sheet;
  sandbox.clearNewsByTreatment_ = id => cleared.push(id);
  sandbox.invalidatePatientCaches_ = (pid, opts) => invalidated.push({ pid, opts });
  sandbox.listTreatmentsForCurrentMonth = pid => { fetched.push(pid); return [`list-${pid}`]; };
  sandbox.log_ = () => {};
  sandbox.Session = { getScriptTimeZone: () => 'Asia/Tokyo' };
  sandbox.Utilities = { formatDate: (date) => formatYearMonth(date) };

  const result = sandbox.deleteTreatmentRow('tid-b');

  assert.strictEqual(sheet.getLastRow(), 3, '2 data rows + header should remain');
  assert.deepStrictEqual(sheet.getRange(2, 7, 2, 1).getValues().flat(), ['tid-a', 'tid-c']);
  assert.deepStrictEqual(cleared, ['tid-b']);
  assert.deepStrictEqual(fetched, ['1']);
  assert.strictEqual(result.patientId, '1');
  assert.ok(Array.isArray(result.treatments), 'updated treatments should be returned');
  assert.strictEqual(invalidated.length, 1);
  assert.strictEqual(invalidated[0].pid, '0001');
  assert.strictEqual(invalidated[0].opts.header, true);
  assert.strictEqual(invalidated[0].opts.treatments, true);
  assert.strictEqual(invalidated[0].opts.latestTreatmentRow, true);
}

function testDeleteRequiresTreatmentId() {
  const now = new Date();
  const currentDate = new Date(now.getFullYear(), now.getMonth(), 1, 9, 0, 0);
  const sheet = createSheet([
    ['TS', '患者ID', '所見', 'email', '', '', '施術ID'],
    [currentDate, '0001', 'first', 'a@example.com', '', '', 'tid-a']
  ]);

  sandbox.sh = () => sheet;
  sandbox.clearNewsByTreatment_ = () => {};
  sandbox.invalidatePatientCaches_ = () => {};
  sandbox.listTreatmentsForCurrentMonth = () => [];
  sandbox.log_ = () => {};
  sandbox.Session = { getScriptTimeZone: () => 'Asia/Tokyo' };
  sandbox.Utilities = { formatDate: (date) => formatYearMonth(date) };

  assert.throws(() => sandbox.deleteTreatmentRow(''), /施術ID/);
}


function testDeleteRejectsPreviousMonth() {
  const now = new Date();
  const previousMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 15, 9, 0, 0);
  const sheet = createSheet([
    ['TS', '患者ID', '所見', 'email', '', '', '施術ID'],
    [previousMonthDate, '0001', 'first', 'a@example.com', '', '', 'tid-prev']
  ]);

  sandbox.sh = () => sheet;
  sandbox.clearNewsByTreatment_ = () => {};
  sandbox.invalidatePatientCaches_ = () => {};
  sandbox.listTreatmentsForCurrentMonth = () => [];
  sandbox.log_ = () => {};
  sandbox.Session = { getScriptTimeZone: () => 'Asia/Tokyo' };
  sandbox.Utilities = { formatDate: (date) => formatYearMonth(date) };

  assert.throws(() => sandbox.deleteTreatmentRow('tid-prev'), /当月以外の施術記録は削除できません/);
}

testDeleteUsesTreatmentId();
testDeleteRequiresTreatmentId();
testDeleteRejectsPreviousMonth();

console.log('deleteTreatmentRow tests passed');
