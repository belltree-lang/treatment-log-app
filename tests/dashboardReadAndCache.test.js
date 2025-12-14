const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

function loadDashboardScripts(ctx, files) {
  files.forEach(file => {
    const code = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', file), 'utf8');
    vm.runInContext(code, ctx);
  });
}

function createBaseContext() {
  const cacheStore = {};
  const propertyStore = {
    data: {},
    getProperty(key) { return this.data[key] || null; },
    setProperty(key, value) { this.data[key] = value; }
  };

  const ctx = {
    console,
    CacheService: {
      getScriptCache: () => ({
        get: key => cacheStore[key] || null,
        put: (key, value) => { cacheStore[key] = value; },
        remove: key => { delete cacheStore[key]; }
      })
    },
    PropertiesService: {
      getScriptProperties: () => propertyStore
    },
    cacheStore,
    propertyStore,
    dashboardWarn_: () => {},
    dashboardResolveTimeZone_: () => 'Asia/Tokyo'
  };

  vm.createContext(ctx);
  return ctx;
}

function createNotesContext() {
  const ctx = createBaseContext();
  const values = [
    [new Date('2025-01-02T09:00:00Z'), 'P001', 'author1', 'latest note'],
    [new Date('2024-12-31T00:00:00Z'), 'P002', 'author2', 'old note']
  ];
  const displays = values.map(row => row.map(cell => cell instanceof Date ? cell.toISOString() : cell));
  let rangeCalls = 0;
  const sheet = {
    getLastRow: () => values.length + 1,
    getLastColumn: () => 5,
    getRange: () => {
      rangeCalls++;
      return {
        getValues: () => values,
        getDisplayValues: () => displays
      };
    }
  };
  ctx.dashboardGetSpreadsheet_ = () => ({ getSheetByName: () => sheet });
  ctx.rangeCalls = () => rangeCalls;
  loadDashboardScripts(ctx, ['cacheUtils.js', 'loadNotes.js', 'markAsRead.js']);
  return ctx;
}

function testMarkAsReadStoresPerUser() {
  const ctx = createBaseContext();
  loadDashboardScripts(ctx, ['cacheUtils.js', 'loadNotes.js', 'markAsRead.js']);
  const ts = new Date('2025-01-01T00:00:00Z');

  const res = ctx.markAsRead('P001', 'User@example.com', ts);
  assert.strictEqual(res.ok, true);
  const stored = JSON.parse(ctx.propertyStore.getProperty('HANDOVER_LAST_READ'));
  assert.strictEqual(stored.P001['user@example.com'], ts.toISOString());
}

function testLoadNotesUsesCacheAndUpdatesUnreadByUser() {
  const ctx = createNotesContext();
  ctx.propertyStore.setProperty('HANDOVER_LAST_READ', JSON.stringify({
    P001: { 'user@example.com': '2025-01-02T12:00:00.000Z' }
  }));

  const first = ctx.loadNotes({ email: 'user@example.com' });
  assert.strictEqual(ctx.rangeCalls(), 2, 'should read sheet once (values + display)');
  assert.strictEqual(first.notes.P001.unread, false, 'latest note is older than last read');

  // update lastRead to older timestamp -> unread should flip, without additional sheet reads
  ctx.propertyStore.setProperty('HANDOVER_LAST_READ', JSON.stringify({
    P001: { 'user@example.com': '2024-12-30T00:00:00.000Z' }
  }));

  const second = ctx.loadNotes({ email: 'user@example.com' });
  assert.strictEqual(ctx.rangeCalls(), 2, 'cache should prevent extra sheet reads');
  assert.strictEqual(second.notes.P001.unread, true, 'note should become unread when lastRead is older');
}

testMarkAsReadStoresPerUser();
testLoadNotesUsesCacheAndUpdatesUnreadByUser();

console.log('dashboard read/cache tests passed');
