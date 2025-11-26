const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const billingGetCode = fs.readFileSync(path.join(__dirname, '../src/get/billingGet.js'), 'utf8');

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
  columnNumberToLetter_: () => ''
};

vm.createContext(context);
vm.runInContext(billingGetCode, context);

const { normalizeBurdenRateInt_ } = context;

if (typeof normalizeBurdenRateInt_ !== 'function') {
  throw new Error('normalizeBurdenRateInt_ failed to load in the test context');
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

function run() {
  testFullWidthDigitsAreParsed();
  testAsciiInputsRemainCompatible();
  console.log('billingGet burden rate tests passed');
}

run();
