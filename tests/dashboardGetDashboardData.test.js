const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const sheetUtilsCode = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard', 'utils', 'sheetUtils.js'),
  'utf8'
);
const configCode = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard', 'config.gs'),
  'utf8'
);
const dashboardCode = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard', 'api', 'getDashboardData.js'),
  'utf8'
);

function createContext(overrides = {}) {
  const context = {
    console,
    JSON,
    Date,
    Set,
    Utilities: {
      formatDate: (date, _tz, _fmt) => date.toISOString()
    },
    Session: {
      getScriptTimeZone: () => 'Asia/Tokyo',
      getActiveUser: () => ({ getEmail: () => overrides.activeEmail || 'session@example.com' })
    }
  };
  Object.assign(context, overrides);
  vm.createContext(context);
  vm.runInContext(configCode, context);
  vm.runInContext(sheetUtilsCode, context);
  vm.runInContext(dashboardCode, context);
  return context;
}

function testAggregatesDashboardData() {
  const patientInfo = {
    patients: { '001': { name: '山田太郎', consentExpiry: '2025-01-31' } },
    warnings: ['p1']
  };
  const notes = {
    notes: { '001': { preview: '最新メモ', when: '2025-02-01', unread: true, authorEmail: 'note@example.com', lastReadAt: '2025-01-01T00:00:00Z', row: 5 } },
    warnings: ['n1']
  };
  const aiReports = { reports: { '001': '2025-01-15 10:00' }, warnings: ['a1'] };
  const invoices = { invoices: { '001': 'https://example.com/invoice.pdf' }, warnings: ['i1'] };
  const treatmentLogs = { logs: [], warnings: ['t1'] };
  const responsible = { responsible: { '001': 'staff@example.com' }, warnings: ['r1'] };
  const tasksResult = { tasks: [{ type: 'consentWarning', patientId: '001' }], warnings: ['task'] };
  const visitsResult = { visits: [{ patientId: '001', time: '10:00' }], warnings: ['visit'] };

  const ctx = createContext();
  const result = ctx.getDashboardData({
    user: 'user@example.com',
    patientInfo,
    notes,
    aiReports,
    invoices,
    treatmentLogs,
    responsible,
    tasksResult,
    visitsResult
  });

  assert.strictEqual(result.meta.user, 'user@example.com');
  assert.ok(result.meta.generatedAt, 'generatedAt should be present');
  assert.deepStrictEqual(result.tasks, tasksResult.tasks);
  assert.deepStrictEqual(result.todayVisits, visitsResult.visits);
  assert.strictEqual(result.patients.length, 1);
  const normalizedPatient = JSON.parse(JSON.stringify(result.patients[0]));
  assert.deepStrictEqual(normalizedPatient, {
    patientId: '001',
    name: '山田太郎',
    consentExpiry: '2025-01-31',
    responsible: 'staff@example.com',
    invoiceUrl: 'https://example.com/invoice.pdf',
    aiReportAt: '2025-01-15 10:00',
    note: {
      patientId: '001',
      preview: '最新メモ',
      when: '2025-02-01',
      unread: true,
      lastReadAt: '2025-01-01T00:00:00Z',
      authorEmail: 'note@example.com',
      row: 5
    }
  });
  const warnings = JSON.parse(JSON.stringify(result.warnings)).sort();
  assert.deepStrictEqual(warnings, ['a1', 'i1', 'n1', 'p1', 'r1', 't1', 'task', 'visit'].sort());
}

function testErrorIsCapturedInMeta() {
  const ctx = createContext({
    loadPatientInfo: () => { throw new Error('boom'); }
  });

  const result = ctx.getDashboardData();
  assert.strictEqual(result.tasks.length, 0);
  assert.strictEqual(result.todayVisits.length, 0);
  assert.strictEqual(result.patients.length, 0);
  assert.ok(result.meta.error && result.meta.error.indexOf('boom') >= 0);
}

(function run() {
  testAggregatesDashboardData();
  testErrorIsCapturedInMeta();
  console.log('dashboardGetDashboardData tests passed');
})();
