const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');
const sandbox = { console };
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
  const sheet = createSheet([
    ['TS', '患者ID', '所見', 'email', '', '', '施術ID'],
    [new Date('2025-02-01T00:00:00Z'), '0001', 'first', 'a@example.com', '', '', 'tid-a'],
    [new Date('2025-02-02T00:00:00Z'), '0001', 'second', '', '', '', 'tid-b'],
    [new Date('2025-02-03T00:00:00Z'), '0001', 'third', '', '', '', 'tid-c']
  ]);

  const cleared = [];
  const invalidated = [];
  const fetched = [];

  sandbox.sh = () => sheet;
  sandbox.clearNewsByTreatment_ = id => cleared.push(id);
  sandbox.invalidatePatientCaches_ = (pid, opts) => invalidated.push({ pid, opts });
  sandbox.listTreatmentsForCurrentMonth = pid => { fetched.push(pid); return [`list-${pid}`]; };
  sandbox.log_ = () => {};

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
  const sheet = createSheet([
    ['TS', '患者ID', '所見', 'email', '', '', '施術ID'],
    [new Date('2025-02-01T00:00:00Z'), '0001', 'first', 'a@example.com', '', '', 'tid-a']
  ]);

  sandbox.sh = () => sheet;
  sandbox.clearNewsByTreatment_ = () => {};
  sandbox.invalidatePatientCaches_ = () => {};
  sandbox.listTreatmentsForCurrentMonth = () => [];
  sandbox.log_ = () => {};

  assert.throws(() => sandbox.deleteTreatmentRow(''), /施術ID/);
}

testDeleteUsesTreatmentId();
testDeleteRequiresTreatmentId();

console.log('deleteTreatmentRow tests passed');
