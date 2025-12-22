const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const mainCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'main.gs'), 'utf8');

function createSheet({ headers, rows }) {
  const data = [headers.slice()].concat(rows.map(row => row.slice()));
  let lastValidation = null;

  return {
    name: '',
    getLastRow: () => data.length,
    getMaxRows: () => data.length,
    getLastColumn: () => Math.max(...data.map(row => row.length)),
    getRange(row, col, numRows, numCols) {
      const zeroRow = row - 1;
      const zeroCol = col - 1;
      const slice = [];
      for (let r = 0; r < numRows; r++) {
        const srcRow = data[zeroRow + r] || [];
        const outRow = [];
        for (let c = 0; c < numCols; c++) {
          outRow.push(srcRow[zeroCol + c] ?? '');
        }
        slice.push(outRow);
      }

      return {
        insertCheckboxes: () => {},
        setDataValidation: (rule) => { lastValidation = rule; },
        getDisplayValues: () => slice.map(r => r.map(v => String(v))),
        getValues: () => slice.map(r => r.slice()),
        setValues(values) {
          for (let r = 0; r < numRows; r++) {
            const targetRow = data[zeroRow + r] || (data[zeroRow + r] = []);
            for (let c = 0; c < numCols; c++) {
              targetRow[zeroCol + c] = values[r][c];
            }
          }
        },
        setValue(value) {
          this.setValues([[value]]);
        }
      };
    },
    insertColumnAfter(index) {
      const insertAt = index;
      for (let i = 0; i < data.length; i++) {
        data[i].splice(insertAt, 0, '');
      }
    },
    insertColumnBefore(index) {
      const insertAt = Math.max(index - 1, 0);
      for (let i = 0; i < data.length; i++) {
        data[i].splice(insertAt, 0, '');
      }
    },
    getLastValidation: () => lastValidation
  };
}

function createContext() {
  const ctx = { console: { warn: () => {}, log: () => {} } };
  vm.createContext(ctx);
  vm.runInContext(mainCode, ctx);

  const validationRule = () => {
    const rule = {};
    return {
      requireCheckbox() { rule.type = 'checkbox'; return this; },
      requireCustomFormulaSatisfied(formula) { rule.formula = formula; return this; },
      setAllowInvalid(allow) { rule.allowInvalid = allow; return this; },
      setHelpText(text) { rule.helpText = text; return this; },
      build() { return Object.assign({}, rule); }
    };
  };

  ctx.SpreadsheetApp = {
    newDataValidation: validationRule
  };

  ctx.resolveBillingColumn_ = (headers, labels, fallbackLabel, options = {}) => {
    const candidates = Array.isArray(labels) ? labels.map(String) : [String(labels || '')];
    for (let i = 0; i < headers.length; i++) {
      const text = headers[i] ? String(headers[i]).trim() : '';
      if (candidates.includes(text)) return i + 1;
    }
    if (options.fallbackLetter) return 1;
    if (options.fallbackIndex) return options.fallbackIndex;
    return options.required ? 1 : 0;
  };

  return ctx;
}

(function testAggregateColumnInsertedNextToUnpaid() {
  const context = createContext();
  const sheet = createSheet({
    headers: ['名前', '未回収チェック', '金額'],
    rows: [
      ['山田太郎', true, 1000],
      ['佐藤花子', false, 2000]
    ]
  });

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const unpaidCol = context.ensureUnpaidCheckColumn_(sheet, headers);
  const aggregateCol = context.ensureAggregateCheckColumn_(sheet, headers, unpaidCol);
  const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

  assert.strictEqual(aggregateCol, unpaidCol + 1, '合算列は未回収列の直後に挿入される');
  assert.strictEqual(updatedHeaders[aggregateCol - 1], '合算', '合算ヘッダーが設定される');
})();

(function testAggregateUnchecksWhenUnpaidIsOff() {
  const context = createContext();
  const sheet = createSheet({
    headers: ['名前', '未回収チェック', '合算', '金額'],
    rows: [
      ['山田太郎', true, true, 1000],
      ['佐藤花子', false, true, 2000]
    ]
  });

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const unpaidCol = context.ensureUnpaidCheckColumn_(sheet, headers);
  const aggregateCol = context.ensureAggregateCheckColumn_(sheet, headers, unpaidCol);

  const aggregateValues = sheet.getRange(2, aggregateCol, 2, 1).getValues();
  assert.deepStrictEqual(aggregateValues, [[true], [false]], '未回収OFFの行は合算が自動で外れる');
})();

(function testAggregateValidationPreventsEnablingWithoutUnpaid() {
  const context = createContext();
  const sheet = createSheet({
    headers: ['名前', '未回収チェック', '合算', '金額'],
    rows: [
      ['山田太郎', '', '', 1000],
      ['佐藤花子', true, '', 2000]
    ]
  });

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const unpaidCol = context.ensureUnpaidCheckColumn_(sheet, headers);
  const aggregateCol = context.ensureAggregateCheckColumn_(sheet, headers, unpaidCol);

  assert.strictEqual(aggregateCol, unpaidCol + 1, '合算列は未回収列の直後');
  const validation = sheet.getLastValidation();
  assert.ok(validation, 'データバリデーションが設定される');
  assert.strictEqual(validation.allowInvalid, false, '未回収OFFでは無効な合算が拒否される');
  assert.strictEqual(validation.formula, '=OR(NOT(C2), $B2)', '未回収列を参照するバリデーション式になる');
  assert.ok(validation.helpText.includes('未回収チェックがON'), 'ヘルプテキストで依存関係を案内する');
})();

console.log('billing bank sheet aggregate column tests passed');
