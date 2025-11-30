const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const billingGetCode = fs.readFileSync(path.join(__dirname, '../src/get/billingGet.js'), 'utf8');

let workbook = null;

const context = {
  columnLetterToNumber_: letter => {
    if (!letter) return 0;
    const normalized = String(letter).toUpperCase();
    let result = 0;
    for (let i = 0; i < normalized.length; i++) {
      const code = normalized.charCodeAt(i);
      if (code < 65 || code > 90) continue;
      result = result * 26 + (code - 64);
    }
    return result;
  },
  columnNumberToLetter_: () => '',
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: () => ''
    })
  },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => workbook
  },
  Session: {
    getScriptTimeZone: () => 'Asia/Tokyo'
  },
  Utilities: {
    formatDate: (date, tz, format) => {
      if (format === 'yyyyMM') {
        const year = date.getFullYear();
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        return `${year}${month}`;
      }
      return '';
    }
  }
};

vm.createContext(context);
vm.runInContext(billingGetCode, context);

const { normalizeBurdenRateInt_ } = context;
const { extractUnpaidBillingHistory } = context;

if (typeof normalizeBurdenRateInt_ !== 'function') {
  throw new Error('normalizeBurdenRateInt_ failed to load in the test context');
}
if (typeof extractUnpaidBillingHistory !== 'function') {
  throw new Error('extractUnpaidBillingHistory failed to load in the test context');
}

function testFullWidthDigitsAreParsed() {
  assert.strictEqual(normalizeBurdenRateInt_('３割'), 3, '全角1〜3割は整数に変換される');
  assert.strictEqual(normalizeBurdenRateInt_('３'), 3, '全角数字だけの場合も整数に変換される');
  assert.strictEqual(normalizeBurdenRateInt_('３０％'), 3, '全角パーセント表記も正しく解釈される');
}

function testAsciiInputsRemainCompatible() {
  assert.strictEqual(normalizeBurdenRateInt_('2割'), 2, '従来の半角入力も維持される');
  assert.strictEqual(normalizeBurdenRateInt_(0.2), 2, '数値入力もそのまま正規化される');
}

function testPercentageInputsAreRounded() {
  assert.strictEqual(normalizeBurdenRateInt_('25%'), 3, '百分率入力は10%刻みに丸められる');
  assert.strictEqual(normalizeBurdenRateInt_(' ５０ ％ '), 5, '全角・スペース混在でもパーセントを解釈する');
}

function testExtractUnpaidBillingHistory() {
  const sheetValues = [
    ['202501', '001', '山田太郎', 1000, 0, 1000, 500, 500, 'OK', new Date('2025-02-01'), ''],
    ['202412', '002', '佐藤花子', 2000, 0, 2000, 0, 2000, '', '', ''],
    ['202411', '003', '田中一郎', 3000, 0, 3000, 3000, 0, '', '', '']
  ];

  const historySheet = {
    getLastRow: () => sheetValues.length + 1,
    getRange: () => ({ getValues: () => sheetValues })
  };

  workbook = {
    getSheetByName: name => (name === '請求履歴' ? historySheet : null)
  };

  const entries = extractUnpaidBillingHistory('202502');
  assert.strictEqual(entries.length, 2, '対象月以前の未回収のみ抽出される');
  assert.deepStrictEqual(entries.map(e => e.billingMonth), ['202501', '202412'], '古い請求月順に保持される');
  assert.strictEqual(entries[0].unpaidAmount, 500, '未回収額が保持される');
}

function run() {
  testFullWidthDigitsAreParsed();
  testAsciiInputsRemainCompatible();
  testPercentageInputsAreRounded();
  testExtractUnpaidBillingHistory();
  console.log('billingGet burden rate tests passed');
}

run();
