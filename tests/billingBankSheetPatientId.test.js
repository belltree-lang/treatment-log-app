const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const mainCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'main.gs'), 'utf8');

function createSheet({ headers, rows }) {
  const data = [headers.slice()].concat(rows.map(row => row.slice()));

  return {
    name: '',
    setName(value) { this.name = value; },
    getName() { return this.name; },
    getLastRow: () => data.length,
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
    copyTo(workbook) {
      const clone = createSheet({ headers, rows });
      workbook._sheets.push(clone);
      return clone;
    }
  };
}

function createContext() {
  const ctx = { console: { warn: () => {}, log: () => {} } };
  vm.createContext(ctx);
  vm.runInContext(mainCode, ctx);

  ctx.normalizeBillingMonthInput = value => ({
    key: typeof value === 'string' ? value : '202501',
    year: 2025,
    month: 1
  });
  ctx.billingNormalizePatientId_ = value => (value ? String(value).trim() : '');
  ctx.columnLetterToNumber_ = letter => {
    if (!letter) return 0;
    let num = 0;
    const chars = String(letter).toUpperCase();
    for (let i = 0; i < chars.length; i++) {
      num = num * 26 + (chars.charCodeAt(i) - 64);
    }
    return num;
  };
  ctx.resolveBillingColumn_ = (headers, labels, fallbackLabel, options = {}) => {
    const candidates = Array.isArray(labels) ? labels.map(String) : [String(labels || '')];
    for (let i = 0; i < headers.length; i++) {
      const text = headers[i] ? String(headers[i]).trim() : '';
      if (candidates.includes(text)) return i + 1;
    }
    if (options.fallbackLetter) return ctx.columnLetterToNumber_(options.fallbackLetter);
    if (options.fallbackIndex) return options.fallbackIndex;
    return options.required ? 1 : 0;
  };
  ctx.BILLING_LABELS = { name: ['名前'], furigana: ['フリガナ'], recNo: ['患者ID'] };

  return ctx;
}

(function testPatientIdColumnOverridesNameMatching() {
  const context = createContext();
  const template = createSheet({
    headers: ['名前', 'フリガナ', '患者ID', '金額'],
    rows: [
      ['山田太郎', 'やまだたろう', 'P001', ''],
      ['山田太郎', 'やまだたろう', 'P002', '']
    ]
  });

  const workbook = {
    _sheets: [],
    getSheetByName: name => workbook._sheets.find(s => s.getName && s.getName() === name) || null,
    deleteSheet: sheet => { workbook._sheets = workbook._sheets.filter(s => s !== sheet); },
    setActiveSheet: () => {},
    moveActiveSheet: () => {},
    getNumSheets: () => workbook._sheets.length,
    billingLogger_: null
  };

  context.billingLogger_ = { log: () => {} };
  context.billingSs = () => workbook;
  context.ensureBankInfoSheet_ = () => template;
  context.formatBankWithdrawalSheetName_ = () => '銀行引落_202501';
  context.prepareBillingData = () => ({
    billingMonth: '202501',
    patients: {
      P001: { nameKanji: '山田太郎', nameKana: 'やまだたろう' },
      P002: { nameKanji: '山田太郎', nameKana: 'やまだたろう' }
    },
    billingJson: [
      { patientId: 'P001', grandTotal: 1000 },
      { patientId: 'P002', grandTotal: 2000 }
    ]
  });

  const result = context.generateSimpleBankSheet('202501');

  const created = workbook._sheets.find(s => s.getName && s.getName() === '銀行引落_202501');
  const amountRange = created.getRange(2, 4, 2, 1).getValues();

  assert.deepStrictEqual(amountRange, [[1000], [2000]], '患者ID列を優先して金額を補完する');
  assert.strictEqual(result.filled, 2, '2件の引落額が埋まる');
})();

(function testPatientIdColumnIsAddedWhenMissing() {
  const context = createContext();
  const template = createSheet({
    headers: ['名前', 'フリガナ', '金額'],
    rows: [
      ['山田花子', 'やまだはなこ', '']
    ]
  });

  const workbook = {
    _sheets: [],
    getSheetByName: name => workbook._sheets.find(s => s.getName && s.getName() === name) || null,
    deleteSheet: sheet => { workbook._sheets = workbook._sheets.filter(s => s !== sheet); },
    setActiveSheet: () => {},
    moveActiveSheet: () => {},
    getNumSheets: () => workbook._sheets.length,
    billingLogger_: null
  };

  context.billingLogger_ = { log: () => {} };
  context.billingSs = () => workbook;
  context.ensureBankInfoSheet_ = () => template;
  context.formatBankWithdrawalSheetName_ = () => '銀行引落_202501';
  context.prepareBillingData = () => ({
    billingMonth: '202501',
    patients: {
      P003: { nameKanji: '山田花子', nameKana: 'やまだはなこ' }
    },
    billingJson: [
      { patientId: 'P003', grandTotal: 3000 }
    ]
  });

  const result = context.generateSimpleBankSheet('202501');
  const created = workbook._sheets.find(s => s.getName && s.getName() === '銀行引落_202501');
  const headers = created.getRange(1, 1, 1, created.getLastColumn()).getDisplayValues()[0];

  assert(headers.includes('患者ID'), '患者ID列がテンプレートに追加される');
  const pidIndex = headers.indexOf('患者ID');
  const amountIndex = headers.indexOf('金額');

  const pidCell = created.getRange(2, pidIndex + 1, 1, 1).getValues();
  const amountCell = created.getRange(2, amountIndex + 1, 1, 1).getValues();

  assert.strictEqual(pidCell[0][0], '', '新設列の患者IDセルは空のまま');
  assert.deepStrictEqual(amountCell, [[3000]], '氏名のみでも金額が補完される');
  assert.strictEqual(result.filled, 1, '1件の引落額が埋まる');
})();

console.log('billing bank sheet patientId tests passed');
