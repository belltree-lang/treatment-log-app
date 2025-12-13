const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const loadInvoicesCode = fs.readFileSync(path.join(__dirname, '../src/dashboard/loadInvoices.js'), 'utf8');

function createIterator(items) {
  let idx = 0;
  return {
    hasNext: () => idx < items.length,
    next: () => items[idx++]
  };
}

function createFile(name, url, updatedAt) {
  return {
    getName: () => name,
    getUrl: () => url,
    getLastUpdated: () => updatedAt
  };
}

function createFolder(name, files = [], subFolders = []) {
  return {
    getName: () => name,
    getFiles: () => createIterator(files),
    getFolders: () => createIterator(subFolders)
  };
}

function createContext() {
  const Utilities = {
    formatDate: (date, tz, format) => {
      const iso = date.toISOString();
      if (format === 'yyyy-MM') return iso.slice(0, 7);
      if (format === 'yyyy年MM月') {
        const y = iso.slice(0, 4);
        const m = iso.slice(5, 7);
        return `${y}年${m}月`;
      }
      return iso.replace(/-/g, '').slice(0, 6);
    }
  };
  const context = { console, Utilities, Session: { getScriptTimeZone: () => 'Asia/Tokyo' } };
  vm.createContext(context);
  vm.runInContext(loadInvoicesCode, context);
  return context;
}

function testLoadsInvoiceLinksForCurrentMonth() {
  const now = new Date('2025-03-15T00:00:00Z');
  const files = [
    createFile('山田太郎_2025-03_請求書.pdf', 'https://example.com/yamada-old', new Date('2025-03-05T00:00:00Z')),
    createFile('山田太郎_20250312_請求書.pdf', 'https://example.com/yamada-new', new Date('2025-03-12T00:00:00Z')),
    createFile('田中花子_2025-03_請求書.pdf', 'https://example.com/tanaka', new Date('2025-03-08T00:00:00Z')),
    createFile('山田太郎_2024-12_請求書.pdf', 'https://example.com/old', new Date('2024-12-01T00:00:00Z'))
  ];
  const targetFolder = createFolder('202503請求書_山田', files);
  const otherFolder = createFolder('202402請求書_佐藤', [
    createFile('佐藤次郎_2024-02_請求書.pdf', 'https://example.com/sato', new Date('2024-02-10T00:00:00Z'))
  ]);
  const rootFolder = createFolder('root', [], [targetFolder, otherFolder]);

  const patientInfo = {
    patients: {
      '001': { name: '山田太郎' },
      '002': { name: '田中花子' },
      '003': { name: '未作成患者' }
    },
    nameToId: {
      '山田太郎': '001',
      '田中花子': '002'
    },
    warnings: []
  };

  const ctx = createContext();
  const result = ctx.loadInvoices({ patientInfo, rootFolder, now });
  const invoices = Object.assign({}, result.invoices);

  assert.deepStrictEqual(invoices, {
    '001': 'https://example.com/yamada-new',
    '002': 'https://example.com/tanaka',
    '003': null
  }, '当月フォルダのPDFを患者IDに紐付け、最新のものを選択する');
  assert.strictEqual(result.warnings.length, 0, '警告は発生しない');
}

(function run() {
  testLoadsInvoiceLinksForCurrentMonth();
  console.log('dashboardLoadInvoices tests passed');
})();
