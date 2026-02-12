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
    getName() { return '施術録'; },
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

console.log('month key write tests passed');
