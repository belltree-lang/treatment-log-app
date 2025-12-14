const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const notesCode = fs.readFileSync(path.join(__dirname, '../src/dashboard/loadNotes.js'), 'utf8');

function createSheet(headers, rows) {
  const data = rows || [];
  return {
    getLastRow: () => 1 + data.length,
    getLastColumn: () => headers.length,
    getRange: (row, col, numRows, numCols) => {
      if (row === 1) {
        return { getDisplayValues: () => [headers.slice(col - 1, col - 1 + numCols)] };
      }
      const slice = data.slice(row - 2, row - 2 + numRows).map(r => {
        const rowData = [];
        for (let i = 0; i < numCols; i++) {
          rowData[i] = r[col - 1 + i];
        }
        return rowData;
      });
      return {
        getValues: () => slice,
        getDisplayValues: () => slice
      };
    }
  };
}

function createContext(sheet, propsStore) {
  const workbook = { getSheetByName: name => (name === '申し送り' ? sheet : null) };
  const store = propsStore || {
    data: {},
    getProperty: key => store.data[key],
    setProperty: (key, value) => { store.data[key] = value; }
  };

  const context = {
    console,
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    Utilities: {
      formatDate: (date, tz, format) => {
        const iso = date.toISOString();
        if (format === 'yyyy-MM-dd HH:mm') return iso.replace('T', ' ').slice(0, 16);
        if (format === 'yyyy-MM-dd') return iso.slice(0, 10);
        return iso;
      }
    },
    PropertiesService: {
      getScriptProperties: () => store
    },
    dashboardGetSpreadsheet_: () => workbook
  };

  vm.createContext(context);
  vm.runInContext(notesCode, context);
  return context;
}

function testLatestNotePerPatientAndPreview() {
  const headers = ['日時', '患者ID', '内容', '画像URL'];
  const rows = [
    [new Date('2025-01-01T09:00:00Z'), '001', 'alice@example.com', '最初の申し送りです'],
    [
      new Date('2025-02-01T10:30:00Z'),
      '001',
      'bob@example.com',
      '新しい申し送りテキストがとても長いのでプレビューを切り詰める必要があります'
    ],
    [new Date('2025-02-02T08:00:00Z'), '002', 'carol@example.com', '別患者のメモ']
  ];
  const sheet = createSheet(headers, rows);
  const store = {
    data: { HANDOVER_LAST_READ: JSON.stringify({ '001': '2025-02-01T10:00:00.000Z' }) },
    getProperty: key => store.data[key],
    setProperty: (key, value) => { store.data[key] = value; }
  };
  const ctx = createContext(sheet, store);
  const result = ctx.loadNotes();

  assert.strictEqual(Object.keys(result.notes).length, 2, '患者ごとに最新1件のみ返す');
  assert.strictEqual(result.notes['001'].authorEmail, 'bob@example.com');
  assert.strictEqual(
    result.notes['001'].note,
    '新しい申し送りテキストがとても長いのでプレビューを切り詰める必要があります'
  );
  assert.strictEqual(result.notes['001'].preview, '新しい申し送りテキストがとても長いのでプ');
  assert.strictEqual(result.notes['001'].when, '2025-02-01 10:30');
  assert.strictEqual(result.notes['001'].lastReadAt, '2025-02-01T10:00:00.000Z');
  assert.strictEqual(result.notes['002'].preview, '別患者のメモ');
}

function testWarningsAndMissingDataHandling() {
  const headers = ['日時', '患者ID', '投稿者', '本文'];
  const rows = [
    [new Date('2025-01-10T00:00:00Z'), '', 'nobody@example.com', ''],
    ['invalid-date', '001', 'nobody@example.com', '本文'],
    [new Date('2025-01-11T00:00:00Z'), '002', 'somebody@example.com', 'OK']
  ];
  const sheet = createSheet(headers, rows);
  const ctx = createContext(sheet);
  const result = ctx.loadNotes();

  assert.ok(result.warnings.some(w => w.includes('患者ID')), '患者ID欠如の警告');
  assert.ok(result.warnings.some(w => w.includes('日時')), '日時解釈失敗の警告');
  assert.deepStrictEqual(Object.keys(result.notes), ['002'], '有効な行のみ残す');
}

function testUpdateLastReadPersistsValue() {
  const ctx = createContext(createSheet(['日時', '患者ID', '投稿者', '本文'], []));
  const ok = ctx.updateHandoverLastRead('003', new Date('2025-03-03T12:00:00Z'));

  assert.ok(ok, '保存処理が成功する');
  const stored = ctx.PropertiesService.getScriptProperties().getProperty('HANDOVER_LAST_READ');
  assert.deepStrictEqual(JSON.parse(stored), { '003': '2025-03-03T12:00:00.000Z' });
}

(function run() {
  testLatestNotePerPatientAndPreview();
  testWarningsAndMissingDataHandling();
  testUpdateLastReadPersistsValue();
  console.log('dashboardLoadNotes tests passed');
})();
