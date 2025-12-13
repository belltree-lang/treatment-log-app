const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const aiReportCode = fs.readFileSync(path.join(__dirname, '../src/dashboard/loadAIReports.js'), 'utf8');

function createSheet(rows) {
  const headers = ['TS', '患者ID', 'memo'];
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

function createAiContext(sheet) {
  const workbook = {
    getSheetByName: name => (name === 'AI報告書' ? sheet : null)
  };
  const Utilities = {
    formatDate: (date, tz, format) => {
      const iso = date.toISOString();
      if (format === 'yyyy-MM-dd HH:mm') return iso.replace('T', ' ').slice(0, 16);
      if (format === 'yyyy-MM-dd') return iso.slice(0, 10);
      return iso;
    }
  };

  const context = { console, Utilities, Session: { getScriptTimeZone: () => 'Asia/Tokyo' }, dashboardGetSpreadsheet_: () => workbook };
  vm.createContext(context);
  vm.runInContext(aiReportCode, context);
  return context;
}

function testLatestReportPerPatientIsReturned() {
  const sheet = createSheet([
    [new Date('2025-02-01T00:00:00Z'), '001', 'old'],
    [new Date('2025-03-05T09:00:00Z'), '001', 'new'],
    [new Date('2025-01-10T12:00:00Z'), '002', 'other'],
  ]);
  const ctx = createAiContext(sheet);
  const result = ctx.loadAIReports();
  const reports = Object.assign({}, result.reports);

  assert.deepStrictEqual(reports, {
    '001': '2025-03-05 09:00',
    '002': '2025-01-10 12:00'
  }, '患者ごとに最新TSが取り出される');
}

function testMissingSheetReturnsWarning() {
  const workbook = { getSheetByName: () => null };
  const context = {
    console,
    Utilities: { formatDate: () => '' },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    dashboardGetSpreadsheet_: () => workbook
  };
  vm.createContext(context);
  vm.runInContext(aiReportCode, context);
  const result = context.loadAIReports();
  const reports = Object.assign({}, result.reports);

  assert.deepStrictEqual(reports, {}, 'シートなしの場合は空オブジェクト');
  assert.ok(result.warnings.some(w => w.includes('AI報告書')), 'シート欠如の警告が含まれる');
}

(function run() {
  testLatestReportPerPatientIsReturned();
  testMissingSheetReturnsWarning();
  console.log('dashboardLoadAIReports tests passed');
})();
