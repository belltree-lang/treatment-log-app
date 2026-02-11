const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(code, sandbox);

function formatYearMonth(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  return `${yyyy}${mm}`;
}

function createSheet(rows) {
  return {
    rows: rows.map(r => r.slice()),
    getLastRow() { return this.rows.length; },
    getRange(row, col) {
      const self = this;
      return {
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
    }
  };
}

function setupCommonStubs(sheet) {
  sandbox.sh = () => sheet;
  sandbox.log_ = () => {};
  sandbox.invalidatePatientCaches_ = () => {};
  sandbox.Session = { getScriptTimeZone: () => 'Asia/Tokyo' };
  sandbox.Utilities = { formatDate: date => formatYearMonth(date) };
}

function testUpdateAllowsCurrentMonth() {
  const now = new Date();
  const currentMonthDate = new Date(now.getFullYear(), now.getMonth(), 10, 9, 0, 0);
  const sheet = createSheet([
    ['TS', '患者ID', '所見', 'email'],
    [currentMonthDate, '0001', 'before', 'a@example.com']
  ]);
  setupCommonStubs(sheet);

  const result = sandbox.updateTreatmentRow(2, 'after');
  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet.getRange(2, 3).getValue(), 'after');
}

function testUpdateRejectsPreviousMonth() {
  const now = new Date();
  const previousMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 10, 9, 0, 0);
  const sheet = createSheet([
    ['TS', '患者ID', '所見', 'email'],
    [previousMonthDate, '0001', 'before', 'a@example.com']
  ]);
  setupCommonStubs(sheet);

  assert.throws(() => sandbox.updateTreatmentRow(2, 'after'), /当月以外の施術記録は編集できません/);
}

testUpdateAllowsCurrentMonth();
testUpdateRejectsPreviousMonth();

console.log('updateTreatmentRow month guard tests passed');
