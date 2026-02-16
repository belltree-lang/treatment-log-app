const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const sheetUtilsCode = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard', 'utils', 'sheetUtils.js'),
  'utf8'
);
const loadPatientInfoCode = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard', 'data', 'loadPatientInfo.js'),
  'utf8'
);

function createContext() {
  const ctx = {
    console,
    Logger: { log: () => {} },
    dashboardWarn_: () => {},
    dashboardNormalizePatientId_: value => String(value == null ? '' : value).trim(),
    dashboardNormalizeNameKey_: value => String(value == null ? '' : value).trim()
  };
  vm.createContext(ctx);
  vm.runInContext(sheetUtilsCode, ctx);
  vm.runInContext(loadPatientInfoCode, ctx);
  return ctx;
}

function createSheet(headers, rows) {
  const all = [headers].concat(rows);
  return {
    getLastRow: () => all.length,
    getLastColumn: () => headers.length,
    getRange(row, col, numRows, numCols) {
      return {
        getValues() {
          return all.slice(row - 1, row - 1 + numRows).map(r => r.slice(col - 1, col - 1 + numCols));
        },
        getDisplayValues() {
          return all.slice(row - 1, row - 1 + numRows).map(r =>
            r.slice(col - 1, col - 1 + numCols).map(cell => String(cell == null ? '' : cell))
          );
        }
      };
    }
  };
}

(function testLoadPatientInfoWithoutConsentExpiryColumnDoesNotThrow() {
  const ctx = createContext();
  const sheet = createSheet(
    ['患者ID', '氏名', '同意年月日', '備考'],
    [
      ['001', '患者A', '2024-08-16', ''],
      ['002', '患者B', '', '']
    ]
  );
  const workbook = { getSheetByName: () => sheet };

  const result = ctx.loadPatientInfoUncached_({ dashboardSpreadsheet: workbook });

  assert.strictEqual(Object.keys(result.patients).length, 2);
  assert.strictEqual(result.patients['001'].raw['同意年月日'], '2024-08-16');
  assert.strictEqual(result.patients['001'].consentExpiry, '');
})();

console.log('dashboard loadPatientInfo tests passed');
