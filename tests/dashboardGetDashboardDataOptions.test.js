const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const sheetUtilsCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'utils', 'sheetUtils.js'), 'utf8');
const configCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'config.gs'), 'utf8');
const dashboardCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'api', 'getDashboardData.js'), 'utf8');

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
      getActiveUser: () => ({ getEmail: () => 'session@example.com' })
    }
  };
  Object.assign(context, overrides);
  vm.createContext(context);
  vm.runInContext(configCode, context);
  vm.runInContext(sheetUtilsCode, context);
  vm.runInContext(dashboardCode, context);
  return context;
}

function testOptionsArePropagated() {
  const now = new Date('2025-04-15T00:00:00Z');
  const seen = {};
  const patientInfo = { patients: {}, nameToId: {}, warnings: [] };
  const workbook = { getSheetByName: () => null };
  const ctx = createContext({
    loadPatientInfo: () => patientInfo,
    loadNotes: () => ({ notes: {}, warnings: [] }),
    loadAIReports: () => ({ reports: {}, warnings: [] }),
    loadInvoices: opts => { seen.invoices = opts; return { invoices: {}, warnings: [] }; },
    loadTreatmentLogs: opts => { seen.treatmentLogs = opts; return { logs: [], warnings: [], lastStaffByPatient: {} }; },
    assignResponsibleStaff: opts => { seen.responsible = opts; return { responsible: {}, warnings: [] }; },
    loadUnpaidAlerts: opts => { seen.unpaidAlerts = opts; return { alerts: [], warnings: [] }; },
    getTasks: () => ({ tasks: [], warnings: [] }),
    getTodayVisits: () => ({ visits: [], warnings: [] })
  });

  ctx.dashboardGetSpreadsheet_ = () => workbook;
  ctx.getDashboardData({ now, cache: false });

  assert.strictEqual(seen.invoices.now, now, 'loadInvoices に now が伝搬する');
  assert.strictEqual(seen.treatmentLogs.dashboardSpreadsheet, workbook, 'loadTreatmentLogs に spreadsheet が伝搬する');
  assert.strictEqual(seen.responsible.dashboardSpreadsheet, workbook, 'assignResponsibleStaff に spreadsheet が伝搬する');
  assert.strictEqual(seen.unpaidAlerts.dashboardSpreadsheet, workbook, 'loadUnpaidAlerts に spreadsheet が伝搬する');
  assert.strictEqual(seen.treatmentLogs.now, now, 'loadTreatmentLogs に now が伝搬する');
  assert.strictEqual(seen.responsible.now, now, 'assignResponsibleStaff に now が伝搬する');
  assert.strictEqual(seen.invoices.cache, false, 'cache:false が請求書読み込みに伝搬する');
  assert.strictEqual(seen.treatmentLogs.cache, false, 'cache:false が施術録読み込みに伝搬する');
  assert.strictEqual(seen.responsible.cache, false, 'cache:false が担当者判定に伝搬する');
  assert.strictEqual(seen.unpaidAlerts.patientInfo, patientInfo, '未回収アラートに患者情報が伝搬する');
  assert.strictEqual(seen.unpaidAlerts.now, now, '未回収アラートに now が伝搬する');
  assert.strictEqual(seen.unpaidAlerts.cache, false, '未回収アラート読み込みも cache:false を受け取る');
}

(function run() {
  testOptionsArePropagated();
  console.log('dashboardGetDashboardData options tests passed');
})();
