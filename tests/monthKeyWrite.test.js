const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(code, sandbox);

function makeSandboxDate(year, monthIndex, day, hour, minute, second) {
  return vm.runInContext(`new Date(${year}, ${monthIndex}, ${day}, ${hour || 0}, ${minute || 0}, ${second || 0})`, sandbox);
}

function createSheet(rows) {
  return {
    rows: rows.map(r => r.slice()),
    getName() { return '施術録'; },
    getLastRow() { return this.rows.length; },
    getRange(row, col, numRows, numCols) {
      const self = this;
      return {
        getValues() {
          const rowCount = Number(numRows) || 1;
          const colCount = Number(numCols) || 1;
          const out = [];
          for (let r = 0; r < rowCount; r += 1) {
            const rowVals = [];
            const sourceRow = self.rows[row - 1 + r] || [];
            for (let c = 0; c < colCount; c += 1) {
              rowVals.push(sourceRow[col - 1 + c]);
            }
            out.push(rowVals);
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
    appendRow(values) {
      this.rows.push(values.slice());
    }
  };
}

function formatDateStub(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const hh = String(date.getHours()).padStart(2, '0');
  const mi = String(date.getMinutes()).padStart(2, '0');
  const ss = String(date.getSeconds()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
}

function setupCommonStubs(sheet) {
  sandbox.assertDomain_ = () => {};
  sandbox.ensureAuxSheets_ = () => {};
  sandbox.sh = () => sheet;
  sandbox.log_ = () => {};
  sandbox.pushNews_ = () => {};
  sandbox.DashboardIndex_updatePatients = () => {};
  sandbox.invalidatePatientCaches_ = () => {};
  sandbox.invalidateTreatmentsCacheForDate_ = () => {};
  sandbox.queueAfterTreatmentJob = () => {};
  sandbox.findTreatmentRowById_ = () => null;
  sandbox.detectRecentDuplicateTreatment_ = () => null;
  sandbox.findExistingTreatmentOnDate_ = () => null;
  sandbox.resolveTreatmentCategoryFromPayload_ = () => ({ key: 'normal', label: '通常', allowEmptyPatientId: false });
  sandbox.resolveTreatmentAttendanceMetrics_ = () => ({ convertedCount: 0, newPatientCount: 0, totalCount: 1 });
  sandbox.LockService = {
    getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} })
  };
  sandbox.Session = {
    getScriptTimeZone: () => 'Asia/Tokyo',
    getActiveUser: () => ({ getEmail: () => 'tester@example.com' })
  };
  sandbox.Utilities = {
    formatDate: formatDateStub,
    getUuid: () => 'uuid-1'
  };
  sandbox.Logger = { log: () => {} };
}

(function testBuildMonthKey() {
  const d = vm.runInContext('new Date(2026, 1, 3, 10, 20, 0)', sandbox);
  assert.strictEqual(sandbox.buildMonthKeyFromDate_(d), '202602');
  assert.strictEqual(sandbox.buildMonthKeyFromDate_('2026-02-03'), '');
})();

(function testSubmitTreatmentWritesMonthKeyToColumnM() {
  const sheet = createSheet([
    ['ts', 'pid', 'note', 'user', '', '', 'tid', 'catLabel', 'c', 'n', 't', 'memo', 'monthKey']
  ]);
  setupCommonStubs(sheet);

  const result = sandbox.submitTreatment({
    patientId: 'P001',
    notesParts: { note: 'テスト' }
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet.getLastRow(), 2);
  const writtenDate = new Date(sheet.rows[1][0]);
  const expectedMonthKey = `${writtenDate.getFullYear()}${String(writtenDate.getMonth() + 1).padStart(2, '0')}`;
  assert.strictEqual(sheet.rows[1][12], expectedMonthKey);
})();

(function testUpdateTreatmentTimestampUpdatesMonthKey() {
  const oldDate = new Date(2025, 0, 31, 9, 0, 0);
  const sheet = createSheet([
    ['ts', 'pid', 'note', 'user', '', '', 'tid', 'catLabel', 'c', 'n', 't', 'memo', 'monthKey'],
    [oldDate, 'P001', 'テスト', 'user@example.com', '', '', 'tid-1', '', '', '', '', '', '202501']
  ]);
  setupCommonStubs(sheet);

  const ok = sandbox.updateTreatmentTimestamp(2, '2025-02-01T10:30');
  assert.strictEqual(ok, true);
  assert.strictEqual(sheet.rows[1][12], '202502');
})();

(function testBackfillMonthKeyWithLimit() {
  const sheet = createSheet([
    ['ts', 'pid', 'note', 'user', '', '', 'tid', 'catLabel', 'c', 'n', 't', 'memo', 'monthKey'],
    [makeSandboxDate(2025, 0, 31, 9, 0, 0), 'P001', '', '', '', '', '', '', '', '', '', '', ''],
    [makeSandboxDate(2025, 1, 1, 10, 0, 0), 'P002', '', '', '', '', '', '', '', '', '', '', ''],
    [makeSandboxDate(2025, 1, 2, 10, 0, 0), 'P003', '', '', '', '', '', '', '', '', '', '', '202502']
  ]);
  setupCommonStubs(sheet);

  const result = sandbox.backfillMonthKey_(1);
  assert.strictEqual(result.updated, 1);
  assert.strictEqual(sheet.rows[1][12], '202501');
  assert.strictEqual(sheet.rows[2][12], '');
  assert.strictEqual(sheet.rows[3][12], '202502');
})();

(function testBackfillMonthKeyWithoutLimit() {
  const sheet = createSheet([
    ['ts', 'pid', 'note', 'user', '', '', 'tid', 'catLabel', 'c', 'n', 't', 'memo', 'monthKey'],
    [makeSandboxDate(2025, 2, 1, 9, 0, 0), 'P001', '', '', '', '', '', '', '', '', '', '', ''],
    ['not-date', 'P002', '', '', '', '', '', '', '', '', '', '', ''],
    [makeSandboxDate(2025, 2, 2, 9, 0, 0), 'P003', '', '', '', '', '', '', '', '', '', '', '']
  ]);
  setupCommonStubs(sheet);

  const result = sandbox.backfillMonthKey_();
  assert.strictEqual(result.updated, 2);
  assert.strictEqual(sheet.rows[1][12], '202503');
  assert.strictEqual(sheet.rows[2][12], '');
  assert.strictEqual(sheet.rows[3][12], '202503');
})();

console.log('month key write tests passed');
