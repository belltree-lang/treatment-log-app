const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');
const logs = [];
const sandbox = {
  console: {
    log: (...args) => logs.push(args.join(' ')),
    warn: () => {},
    error: () => {},
  },
  Utilities: {
    formatDate(date, _tz, fmt) {
      const yyyy = date.getFullYear();
      const mm = String(date.getMonth() + 1).padStart(2, '0');
      const dd = String(date.getDate()).padStart(2, '0');
      if (fmt === 'yyyy-MM-dd') return `${yyyy}-${mm}-${dd}`;
      const hh = String(date.getHours()).padStart(2, '0');
      const mi = String(date.getMinutes()).padStart(2, '0');
      const ss = String(date.getSeconds()).padStart(2, '0');
      return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
    },
  },
};
vm.createContext(sandbox);
vm.runInContext(code, sandbox);

function createSheet(rows) {
  const calls = [];
  return {
    rows: rows.map(r => r.slice()),
    calls,
    getLastRow() { return this.rows.length; },
    getMaxColumns() { return 13; },
    getRange(row, col, numRows, numCols) {
      calls.push({ row, col, numRows, numCols });
      const self = this;
      return {
        getValues() {
          const out = [];
          for (let r = 0; r < numRows; r += 1) {
            const source = self.rows[row - 1 + r] || [];
            const rowOut = [];
            for (let c = 0; c < numCols; c += 1) rowOut.push(source[col - 1 + c]);
            out.push(rowOut);
          }
          return out;
        }
      };
    }
  };
}

(function testDetectRecentDuplicateTreatmentOptimizedWindow() {
  logs.length = 0;
  const base = new Date('2026-03-15T12:00:00+09:00');
  const rows = [['ts', 'pid', 'note', 'mail', '', '', 'tid', '', '', '', '', '', 'monthKey']];
  for (let i = 0; i < 130; i += 1) {
    rows.push([
      new Date(base.getTime() - (130 - i) * 60000),
      'P-OTHER',
      'x',
      '', '', '', `tid-${i}`,
      '', '', '', '', '',
      '202603'
    ]);
  }
  rows[31] = [new Date('2026-03-15T11:59:30+09:00'), 'P001', '重複ノート', '', '', '', 'tid-target', '', '', '', '', '', '202603'];

  const sheet = createSheet(rows);
  const result = sandbox.detectRecentDuplicateTreatment_(sheet, 'P001', '重複ノート', base, 'Asia/Tokyo', '', '202603');

  assert.ok(result);
  assert.strictEqual(result.reason, 'recentContent');
  assert.strictEqual(sheet.calls[0].numRows, 100);
  assert.ok(logs.some(line => line.includes('[perf][submitTreatment] optimizedDuplicate rows=100')));
})();

(function testFindExistingTreatmentOnDateUsesMonthKeyFilter() {
  logs.length = 0;
  const base = new Date('2026-03-20T18:00:00+09:00');
  const rows = [['ts', 'pid', 'note', 'mail', '', '', 'tid', '', '', '', '', '', 'monthKey']];
  for (let i = 0; i < 140; i += 1) {
    rows.push([
      new Date(base.getTime() - (140 - i) * 3600000),
      'P-OTHER',
      'x',
      '', '', '', `tid-${i}`,
      '', '', '', '', '',
      '202603'
    ]);
  }
  rows[80] = [new Date('2026-03-20T09:30:00+09:00'), 'P001', '他月', '', '', '', 'tid-wrong-month', '', '', '', '', '', '202602'];
  rows[90] = [new Date('2026-03-20T10:30:00+09:00'), 'P001', '同日', '', '', '', 'tid-right', '', '', '', '', '', '202603'];

  const sheet = createSheet(rows);
  const result = sandbox.findExistingTreatmentOnDate_(sheet, 'P001', base, 'Asia/Tokyo', '', '202603');

  assert.ok(result);
  assert.strictEqual(result.treatmentId, 'tid-right');
  assert.strictEqual(sheet.calls[0].numRows, 100);
})();

console.log('submit treatment scan optimization tests passed');
