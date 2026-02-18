const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const configCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'config.gs'), 'utf8');
const sheetUtilsCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'utils', 'sheetUtils.js'), 'utf8');
const roleCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'auth', 'role.js'), 'utf8');
const dashboardCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'api', 'getDashboardData.js'), 'utf8');

function createContext(overrides = {}) {
  const cacheStore = {};
  const logs = [];
  const ctx = {
    console,
    JSON,
    Date,
    Set,
    CacheService: {
      getScriptCache: () => ({
        get: key => cacheStore[key] || null,
        put: (key, value) => { cacheStore[key] = value; }
      })
    },
    Logger: {
      log: message => logs.push(String(message))
    },
    Utilities: {
      formatDate: (date, _tz, format) => {
        const d = new Date(date);
        if (format === 'yyyyMM') {
          return `${d.getUTCFullYear()}${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
        }
        if (format === 'yyyy-MM-dd') {
          return d.toISOString().slice(0, 10);
        }
        return d.toISOString();
      },
      getUuid: () => 'mock-uuid'
    },
    Session: {
      getScriptTimeZone: () => 'Asia/Tokyo',
      getActiveUser: () => ({ getEmail: () => 'session@example.com' })
    },
    SpreadsheetApp: {
      openById: id => ({ id })
    },
    DriveApp: {
      getFolderById: id => ({ id })
    },
    __cacheStore: cacheStore,
    __logs: logs,
    ...overrides
  };
  vm.createContext(ctx);
  ctx.module = { exports: {} };
  vm.runInContext(configCode, ctx);
  vm.runInContext(sheetUtilsCode, ctx);
  vm.runInContext(roleCode, ctx);
  vm.runInContext(dashboardCode, ctx);
  return ctx;
}

function testTreatmentLogsCacheByMonthKey() {
  let loadCount = 0;
  const ctx = createContext({
    DASHBOARD_DEBUG_OVERRIDE: true,
    dashboardGetSpreadsheet_: () => ({ getSheetByName: () => null }),
    loadPatientInfo: () => ({ patients: {}, warnings: [] }),
    loadNotes: () => ({ notes: {}, warnings: [] }),
    loadAIReports: () => ({ reports: {}, warnings: [] }),
    loadInvoices: () => ({ invoices: {}, warnings: [] }),
    loadTreatmentLogs: () => {
      loadCount += 1;
      return { logs: [{ patientId: 'P001', timestamp: new Date('2025-02-01T00:00:00Z') }], warnings: [] };
    },
    assignResponsibleStaff: () => ({ responsible: {}, warnings: [] }),
    loadUnpaidAlerts: () => ({ alerts: [], warnings: [] }),
    getTasks: () => ({ tasks: [], warnings: [] }),
    getTodayVisits: () => ({ visits: [], warnings: [] })
  });

  ctx.getDashboardData({ user: 'user1@example.com', now: new Date('2025-02-15T00:00:00Z') });
  ctx.getDashboardData({ user: 'user2@example.com', now: new Date('2025-02-20T00:00:00Z') });

  assert.strictEqual(loadCount, 1, '同一monthKeyではloadTreatmentLogsは1回のみ');
  assert.ok(ctx.__logs.some(line => line.includes('[perf] cacheMiss key=dashboard:treatmentLogs:202502')));
  assert.ok(ctx.__logs.some(line => line.includes('[perf] cacheHit key=dashboard:treatmentLogs:202502')));
}

function testSpreadsheetCacheInvalidatesById() {
  let openCount = 0;
  const ctx = createContext({
    SpreadsheetApp: {
      openById: id => {
        openCount += 1;
        return { id, openCount };
      }
    }
  });

  const first = ctx.dashboardGetSpreadsheet_();
  const second = ctx.dashboardGetSpreadsheet_();
  assert.strictEqual(openCount, 1, '同一IDではopenByIdがキャッシュされる');
  assert.strictEqual(first.id, second.id);

  ctx.dashboardRuntimeCachePut_('dashboard:spreadsheet', { id: 'legacy-sheet-id', spreadsheet: { id: 'legacy-sheet-id' } });
  const third = ctx.dashboardGetSpreadsheet_();
  assert.strictEqual(openCount, 2, 'ID不一致のキャッシュは破棄して再取得する');
  assert.strictEqual(third.id, '1ajnW9Fuvu0YzUUkfTmw0CrbhrM3lM5tt5OA1dK2_CoQ');
  assert.ok(ctx.__logs.some(line => line.includes('[perf] cacheHit key=dashboard:spreadsheet')));
  assert.ok(ctx.__logs.some(line => line.includes('reason=spreadsheetIdChanged')));
}

function testInvoiceRootFolderCache() {
  let folderOpenCount = 0;
  const ctx = createContext({
    DASHBOARD_INVOICE_FOLDER_ID: 'folder-1',
    DriveApp: {
      getFolderById: id => {
        folderOpenCount += 1;
        return { id };
      }
    }
  });

  const first = ctx.dashboardGetInvoiceRootFolder_();
  const second = ctx.dashboardGetInvoiceRootFolder_();
  assert.strictEqual(folderOpenCount, 1, 'invoiceRootFolderをキャッシュする');
  assert.strictEqual(first.id, second.id);
  assert.ok(ctx.__logs.some(line => line.includes('[perf] cacheHit key=dashboard:invoiceRootFolder')));
}

testTreatmentLogsCacheByMonthKey();
testSpreadsheetCacheInvalidatesById();
testInvoiceRootFolderCache();

console.log('dashboard caching perf tests passed');
