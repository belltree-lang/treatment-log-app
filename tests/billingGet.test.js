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
const { loadTreatmentLogs_ } = context;
const { billingParseTreatmentTimestamp_ } = context;
const { loadBillingStaffDirectory_ } = context;

if (typeof normalizeBurdenRateInt_ !== 'function') {
  throw new Error('normalizeBurdenRateInt_ failed to load in the test context');
}
if (typeof extractUnpaidBillingHistory !== 'function') {
  throw new Error('extractUnpaidBillingHistory failed to load in the test context');
}
if (typeof loadTreatmentLogs_ !== 'function') {
  throw new Error('loadTreatmentLogs_ failed to load in the test context');
}
if (typeof billingParseTreatmentTimestamp_ !== 'function') {
  throw new Error('billingParseTreatmentTimestamp_ failed to load in the test context');
}
if (typeof loadBillingStaffDirectory_ !== 'function') {
  throw new Error('loadBillingStaffDirectory_ failed to load in the test context');
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

function testLoadTreatmentLogsDoesNotRequireLogger() {
  const headers = ['日付', '患者ID', '作成者'];
  const dataRows = [
    [new Date('2024-12-01T00:00:00Z'), '001', 'staff@example.com'],
    [new Date('2024-12-02T00:00:00Z'), '002', 'other@example.com']
  ];

  const sheetValues = [headers, ...dataRows];
  const sheetDisplayValues = [
    headers,
    ['2024/12/01 00:00', '001', 'staff@example.com'],
    ['2024/12/02 00:00', '002', 'other@example.com']
  ];

  const sheet = {
    getLastRow: () => sheetValues.length,
    getLastColumn: () => headers.length,
    getMaxColumns: () => headers.length,
    getRange: (row, col, numRows, numCols) => {
      const sliceValues = rows => rows.map(r => r.slice(col - 1, col - 1 + numCols));
      if (row === 1 && numRows === 1) {
        return { getDisplayValues: () => [sheetDisplayValues[0]] };
      }
      return {
        getValues: () => sliceValues(sheetValues.slice(row - 1, row - 1 + numRows)),
        getDisplayValues: () => sliceValues(sheetDisplayValues.slice(row - 1, row - 1 + numRows))
      };
    }
  };

  workbook = {
    getSheetByName: name => (name === '施術録' ? sheet : null)
  };

  const logs = loadTreatmentLogs_();
  assert.strictEqual(logs.length, 2, '施術録が問題なく取得できる');
  assert.strictEqual(Object.prototype.toString.call(logs[0].timestamp), '[object Date]', '日付はDateとして解釈される');
  assert.strictEqual(logs[0].patientId, '001', '患者IDが保持される');
}

function testBillingParseTreatmentTimestampParsesSerialStrings() {
  const serial = '45600.5';
  const parsed = billingParseTreatmentTimestamp_(serial, null);
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const expected = new Date(excelEpoch.getTime() + Math.round(Number(serial) * 24 * 60 * 60 * 1000));
  assert.strictEqual(Object.prototype.toString.call(parsed), '[object Date]', 'シリアル値がDateに変換される');
  assert.strictEqual(parsed.toISOString(), expected.toISOString(), 'シリアル文字列と数値で同じ日付になる');
}

function testBillingParseTreatmentTimestampParsesJapaneseDateText() {
  const parsed = billingParseTreatmentTimestamp_(undefined, '2024年12月01日');
  assert.strictEqual(Object.prototype.toString.call(parsed), '[object Date]', '和暦表記でもDateに変換される');
  assert.strictEqual(parsed.getFullYear(), 2024, '年が正しく解釈される');
  assert.strictEqual(parsed.getMonth(), 11, '月が正しく解釈される（0始まり）');
}

function testStaffDirectoryUsesStaffIdWhenEmailMissing() {
  const headers = ['氏名', 'スタッフID'];
  const values = [
    ['山田太郎', 'STAFF001'],
    ['佐藤花子', 'staff002']
  ];

  const sheet = {
    getLastRow: () => values.length + 1,
    getLastColumn: () => headers.length,
    getMaxColumns: () => headers.length,
    getRange: (row, col, numRows, numCols) => {
      if (row === 1 && numRows === 1) {
        return { getDisplayValues: () => [headers.slice(col - 1, col - 1 + numCols)] };
      }
      const startIdx = Math.max(0, row - 2);
      return {
        getValues: () => values
          .slice(startIdx, startIdx + numRows)
          .map(r => r.slice(col - 1, col - 1 + numCols))
      };
    }
  };

  workbook = {
    getSheetByName: name => (name === 'スタッフ一覧' ? sheet : null)
  };

  const directory = loadBillingStaffDirectory_();
  assert.strictEqual(directory['staff001'], '山田太郎', 'メールなしでも担当者IDで紐づく');
  assert.strictEqual(directory['staff002'], '佐藤花子', '複数行の担当者IDを辞書化する');
}

function run() {
  testFullWidthDigitsAreParsed();
  testAsciiInputsRemainCompatible();
  testPercentageInputsAreRounded();
  testExtractUnpaidBillingHistory();
  testLoadTreatmentLogsDoesNotRequireLogger();
  testBillingParseTreatmentTimestampParsesSerialStrings();
  testBillingParseTreatmentTimestampParsesJapaneseDateText();
  testStaffDirectoryUsesStaffIdWhenEmailMissing();
  console.log('billingGet burden rate tests passed');
}

run();
