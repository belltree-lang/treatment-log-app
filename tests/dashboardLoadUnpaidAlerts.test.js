const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const sheetUtilsCode = fs.readFileSync(path.join(__dirname, '../src/dashboard/utils/sheetUtils.js'), 'utf8');
const configCode = fs.readFileSync(path.join(__dirname, '../src/dashboard/config.gs'), 'utf8');
const cacheUtilsCode = fs.readFileSync(path.join(__dirname, '../src/dashboard/utils/cacheUtils.js'), 'utf8');
const unpaidAlertsCode = fs.readFileSync(path.join(__dirname, '../src/dashboard/data/loadUnpaidAlerts.js'), 'utf8');

function createSheet(rows) {
  const headers = ['患者ID', '対象月', '金額', '理由', '備考', '記録日時'];
  const data = rows || [];
  return {
    getLastRow: () => 1 + data.length,
    getLastColumn: () => headers.length,
    getRange: (row, col, numRows, numCols) => {
      if (row === 1) {
        return { getDisplayValues: () => [headers.slice(col - 1, col - 1 + numCols)] };
      }
      const slice = data.slice(row - 2, row - 2 + numRows).map(r => {
        const rowData = [];
        for (let i = 0; i < numCols; i++) {
          rowData[i] = r[col - 1 + i];
        }
        return rowData;
      });
      return {
        getValues: () => slice,
        getDisplayValues: () => slice
      };
    }
  };
}

function createContext(sheet) {
  const workbook = { getSheetByName: name => (name === '未回収履歴' ? sheet : null) };
  const Utilities = {
    formatDate: (date, _tz, format) => {
      const iso = date.toISOString();
      if (format === 'yyyy-MM') return iso.slice(0, 7);
      if (format === 'yyyy-MM-dd HH:mm') return iso.replace('T', ' ').slice(0, 16);
      return iso;
    }
  };

  const context = {
    console,
    Utilities,
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    dashboardGetSpreadsheet_: () => workbook
  };
  vm.createContext(context);
  vm.runInContext(configCode, context);
  vm.runInContext(sheetUtilsCode, context);
  vm.runInContext(cacheUtilsCode, context);
  context.dashboardGetSpreadsheet_ = () => workbook;
  vm.runInContext(unpaidAlertsCode, context);
  return context;
}

function testConsecutiveUnpaidAlertsAreCollected() {
  const sheet = createSheet([
    ['001', '2024-01-10', 10000, '確認中', '', '2024-02-05T00:00:00Z'],
    ['001', '2023/12/01', 20000, '', '', '2024-01-10T00:00:00Z'],
    ['001', '2023-11-01', 30000, '', '', '2023-12-05T00:00:00Z'],
    ['002', '2023-12-01', 5000, '', '', '2023-12-10T00:00:00Z']
  ]);
  const ctx = createContext(sheet);
  const result = ctx.loadUnpaidAlerts({
    patientInfo: { patients: { '001': { name: '山田太郎' }, '002': { name: '佐藤花子' } }, warnings: [] }
  });

  assert.strictEqual(result.alerts.length, 1, '連続3ヶ月のみが抽出される');
  const alert = result.alerts[0];
  assert.strictEqual(alert.patientId, '001');
  assert.strictEqual(alert.patientName, '山田太郎');
  assert.strictEqual(alert.consecutiveMonths, 3);
  assert.strictEqual(alert.totalAmount, 60000);
  assert.strictEqual(alert.months.map(m => m.key).join(','), '2024-01,2023-12,2023-11');
}

function testMissingSheetProducesWarning() {
  const workbook = { getSheetByName: () => null };
  const context = createContext(null);
  context.dashboardGetSpreadsheet_ = () => workbook;
  const result = context.loadUnpaidAlerts({ patientInfo: { patients: {}, warnings: [] } });

  assert.ok(Array.isArray(result.alerts) && result.alerts.length === 0);
  assert.strictEqual(result.setupIncomplete, true);
  assert.ok(result.warnings.some(w => w.includes('未回収履歴')), 'シート欠如の警告が含まれる');
}

(function run() {
  testConsecutiveUnpaidAlertsAreCollected();
  testMissingSheetProducesWarning();
  console.log('dashboardLoadUnpaidAlerts tests passed');
})();
