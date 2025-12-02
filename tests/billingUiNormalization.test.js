const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/main.js.html'), 'utf8')
  .replace(/^\s*<script>\s*/, '')
  .replace(/\s*<\/script>\s*$/, '');

function createContext() {
  const documentStub = {
    body: { dataset: {} },
    getElementById: () => null,
    querySelector: () => null,
    querySelectorAll: () => []
  };

  const windowStub = { APP_CONFIG: {}, document: documentStub };
  windowStub.window = windowStub;

  const ctx = {
    console,
    alert: () => {},
    document: documentStub,
    window: windowStub,
    Intl,
  };

  vm.createContext(ctx);
  vm.runInContext(code, ctx);
  return ctx;
}

const { normalizeBillingResultPayload } = createContext();

function testParsesStringifiedPayload() {
  const raw = JSON.stringify({ billingJson: [{ patientId: '001' }], billingMonth: '202501' });
  const result = normalizeBillingResultPayload(raw);
  assert.strictEqual(result.billingMonth, '202501', '請求月が保持される');
  assert.strictEqual(result.billingJson.length, 1, 'billingJson が配列として復元される');
}

function testFindsNestedBillingJson() {
  const raw = {
    data: {
      response: {
        billingJson: JSON.stringify([{ patientId: '777' }]),
        preparedAt: '2025-12-01T00:00:00Z'
      }
    }
  };
  const result = normalizeBillingResultPayload(raw);
  assert.strictEqual(result.billingJson[0].patientId, '777', 'ネストした billingJson を抽出する');
  assert.strictEqual(result.preparedAt, '2025-12-01T00:00:00Z', '元のメタデータを保持する');
}

function testReturnsNullOnUnparsableString() {
  const result = normalizeBillingResultPayload('{invalid json');
  assert.strictEqual(result, null, '不正なJSON文字列は null を返す');
}

function run() {
  testParsesStringifiedPayload();
  testFindsNestedBillingJson();
  testReturnsNullOnUnparsableString();
  console.log('billingUiNormalization tests passed');
}

run();
