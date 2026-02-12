const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');

function createNewsSheet(rows) {
  const calls = [];
  return {
    calls,
    getLastRow() { return rows.length; },
    getLastColumn() { return 7; },
    getRange(row, col, numRows, numCols) {
      calls.push({ row, col, numRows, numCols });
      return {
        getValues() {
          const out = [];
          for (let r = 0; r < numRows; r += 1) {
            const src = rows[row - 1 + r] || [];
            const rowVals = [];
            for (let c = 0; c < numCols; c += 1) {
              rowVals.push(src[col - 1 + c] !== undefined ? src[col - 1 + c] : '');
            }
            out.push(rowVals);
          }
          return out;
        },
        getDisplayValues() {
          return this.getValues();
        }
      };
    }
  };
}

(function testGetNewsUsesOptimizedScanAndLimit() {
  const logs = [];
  const cacheWrites = [];
  const cacheStore = new Map();
  const sandbox = {
    console,
    Logger: { log: message => logs.push(String(message)) },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    Utilities: {
      formatDate(date) {
        return date.toISOString();
      }
    },
    CacheService: {
      getScriptCache: () => ({
        get: key => cacheStore.get(key) || null,
        put: (key, value, ttl) => {
          cacheWrites.push({ key, ttl });
          cacheStore.set(key, value);
        }
      })
    }
  };

  vm.createContext(sandbox);
  vm.runInContext(code, sandbox);
  sandbox.standardizeConsentNewsMeta_ = () => {};

  const rows = [['TS', 'pid', 'type', 'message', 'clearedAt', 'meta', 'dismissed']];
  const baseTs = new Date('2026-04-01T00:00:00.000Z').getTime();
  for (let i = 1; i <= 230; i += 1) {
    rows.push([
      new Date(baseTs + i * 60000),
      'P001',
      'info',
      `patient-${i}`,
      '',
      '',
      ''
    ]);
  }
  for (let i = 1; i <= 20; i += 1) {
    rows.push([
      new Date(baseTs + 500000000 + i * 30000),
      '',
      'global',
      `global-${i}`,
      '',
      '',
      ''
    ]);
  }
  for (let i = 1; i <= 20; i += 1) {
    rows.push([
      new Date(baseTs + i * 45000),
      'P999',
      'other',
      `other-${i}`,
      '',
      '',
      ''
    ]);
  }

  const sheet = createNewsSheet(rows);
  sandbox.sh = name => {
    if (name === 'News') return sheet;
    throw new Error(`Unexpected sheet ${name}`);
  };

  const result = sandbox.getNews('P001');
  assert.strictEqual(result.length, 200, '最新200件に制限される');
  assert.strictEqual(result[0].message, 'global-20', '最新データが先頭に来る');
  assert.strictEqual(result[199].message, 'patient-51', '200件目まで保持される');
  assert.ok(result.some(row => row.message === 'global-20'), 'グローバルnewsも返る');

  assert.strictEqual(sheet.calls[0].col, 2, 'pid列を先読みする');
  assert.strictEqual(sheet.calls[0].numCols, 1, 'pid列のみ取得する');
  assert.ok(sheet.calls.slice(1).every(call => call.col === 1 && call.numCols === 7), '一致行は必要列のみ取得する');

  assert.ok(logs.some(line => line.includes('[perf][getNews] pid=P001 totalRows=270 matchedRows=230')));
  assert.ok(cacheWrites.some(write => write.key === 'patient:news:P001' && write.ttl === 180), 'patientキーでTTL180秒キャッシュする');
})();

console.log('getNews performance tests passed');
