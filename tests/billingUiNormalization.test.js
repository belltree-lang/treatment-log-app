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

function testFindsBillingJsonInsideStringifiedPayload() {
  const raw = {
    result: JSON.stringify({
      payload: {
        billingJson: JSON.stringify([{ patientId: '900' }]),
        billingMonth: '202512'
      }
    })
  };

  const result = normalizeBillingResultPayload(raw);
  assert.strictEqual(result.billingJson[0].patientId, '900', '文字列化された payload から billingJson を抽出する');
  assert.strictEqual(result.billingMonth, '202512', 'billingMonth を保持する');
}

function testMergesMetadataFromAncestorObjects() {
  const raw = {
    data: {
      preparedAt: '2025-03-01T00:00:00Z',
      billingMonth: '202503',
      payload: {
        billingJson: JSON.stringify([{ patientId: '555' }])
      }
    }
  };

  const result = normalizeBillingResultPayload(raw);
  assert.strictEqual(result.billingJson[0].patientId, '555', 'ネストした billingJson を抽出する');
  assert.strictEqual(result.billingMonth, '202503', '祖先オブジェクトの billingMonth を引き継ぐ');
  assert.strictEqual(result.preparedAt, '2025-03-01T00:00:00Z', '祖先オブジェクトの preparedAt を引き継ぐ');
}

function testInheritsObjectMetadataFromAncestors() {
  const raw = {
    meta: {
      staffByPatient: { '111': ['taro@example.com'] },
      staffDirectory: { 'taro@example.com': '太郎' },
      carryOverByPatient: { '111': 4000 }
    },
    nested: {
      payload: {
        billingJson: JSON.stringify([{ patientId: '111' }])
      }
    }
  };

  const result = normalizeBillingResultPayload(raw);
  assert.deepStrictEqual(result.staffByPatient['111'], ['taro@example.com'], 'staffByPatient を祖先から引き継ぐ');
  assert.strictEqual(result.staffDirectory['taro@example.com'], '太郎', 'staffDirectory を祖先から引き継ぐ');
  assert.strictEqual(result.carryOverByPatient['111'], 4000, 'carryOverByPatient を祖先から引き継ぐ');
}

function testInheritsFilesFromAncestorMeta() {
  const raw = {
    response: {
      files: [
        { name: 'invoice-001.pdf', url: 'https://example.com/invoice-001.pdf' }
      ],
      payload: {
        billingJson: [{ patientId: '222' }]
      }
    }
  };

  const result = normalizeBillingResultPayload(raw);
  assert.strictEqual(result.billingJson[0].patientId, '222', 'billingJson を抽出する');
  assert.strictEqual(result.files[0].name, 'invoice-001.pdf', '祖先オブジェクトの files 配列を引き継ぐ');
  assert.strictEqual(result.files[0].url, 'https://example.com/invoice-001.pdf', 'files の中身を保持する');
}

function testReturnsNullOnUnparsableString() {
  const result = normalizeBillingResultPayload('{invalid json');
  assert.strictEqual(result, null, '不正なJSON文字列は null を返す');
}

function run() {
  testParsesStringifiedPayload();
  testFindsNestedBillingJson();
  testFindsBillingJsonInsideStringifiedPayload();
  testMergesMetadataFromAncestorObjects();
  testInheritsObjectMetadataFromAncestors();
  testInheritsFilesFromAncestorMeta();
  testReturnsNullOnUnparsableString();
  console.log('billingUiNormalization tests passed');
}

run();
