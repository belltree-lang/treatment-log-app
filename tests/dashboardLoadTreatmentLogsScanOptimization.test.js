const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

function loadDashboardScripts(ctx, files) {
  files.forEach(file => {
    const code = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', file), 'utf8');
    vm.runInContext(code, ctx);
  });
}

function createContext() {
  const logs = [];
  const ctx = {
    console,
    Logger: {
      log: msg => logs.push(String(msg))
    },
    dashboardWarn_: () => {}
  };
  vm.createContext(ctx);
  loadDashboardScripts(ctx, ['utils/sheetUtils.js', 'data/loadTreatmentLogs.js']);
  ctx.__logs = logs;
  return ctx;
}

function createSheet(headers, dataRows) {
  const rows = [headers].concat(dataRows);
  const calls = [];
  return {
    calls,
    getLastRow: () => rows.length,
    getLastColumn: () => headers.length,
    getRange(row, col, numRows, numCols) {
      calls.push({ row, col, numRows, numCols });
      return {
        getValues() {
          const out = [];
          for (let r = 0; r < numRows; r += 1) {
            const src = rows[row - 1 + r] || [];
            const rowOut = [];
            for (let c = 0; c < numCols; c += 1) rowOut.push(src[col - 1 + c]);
            out.push(rowOut);
          }
          return out;
        },
        getDisplayValues() {
          const out = [];
          for (let r = 0; r < numRows; r += 1) {
            const src = rows[row - 1 + r] || [];
            const rowOut = [];
            for (let c = 0; c < numCols; c += 1) {
              const cell = src[col - 1 + c];
              rowOut.push(cell instanceof Date ? cell.toISOString() : String(cell == null ? '' : cell));
            }
            out.push(rowOut);
          }
          return out;
        }
      };
    }
  };
}

(function testLargeSheetTimestampScanStopsEarly() {
  const ctx = createContext();
  const headers = ['日時', '患者ID', '施術者', 'メール', '内容'];
  const rows = [];

  for (let i = 0; i < 5000; i += 1) {
    rows.push([new Date('2024-01-01T09:00:00Z'), 'P-OLD', 'old', 'old@example.com', 'old note']);
  }
  for (let i = 0; i < 120; i += 1) {
    rows.push([
      new Date(Date.UTC(2025, 1, 1, 9, i, 0)),
      i % 2 === 0 ? 'P001' : 'P002',
      'staff',
      'staff@example.com',
      'note'
    ]);
  }

  const sheet = createSheet(headers, rows);
  const patientInfo = {
    nameToId: {},
    warnings: [],
    setupIncomplete: false,
    patients: { P001: { id: 'P001' }, P002: { id: 'P002' } }
  };

  const result = ctx.loadTreatmentLogsUncached_({
    now: new Date('2025-02-15T00:00:00Z'),
    dashboardSpreadsheet: { getSheetByName: () => sheet },
    patientInfo
  });

  assert.ok(result.logs.length > 0, '対象月のログが抽出されること');

  const scanLog = ctx.__logs.find(line => line.includes('[perf] loadTreatmentLogsFilterScan='));
  const matched = scanLog && scanLog.match(/rows=(\d+)\/(\d+)/);
  assert.ok(matched, 'フィルタ範囲探索で行数ログが出力されること');
  const scannedRows = Number(matched[1]);
  const totalRows = Number(matched[2]);
  assert.ok(scannedRows < totalRows, '日時列の全件スキャンを回避すること');

  assert.ok(
    ctx.__logs.some(line => line.includes('[perf] loadTreatmentLogsFilterScan=')),
    'フィルタ範囲探索のperfログが出力されること'
  );
  assert.ok(
    ctx.__logs.some(line => line.includes('[perf] loadTreatmentLogsFilterRows=')),
    'フィルタ後行数のperfログが出力されること'
  );
})();

console.log('dashboard loadTreatmentLogs scan optimization tests passed');
