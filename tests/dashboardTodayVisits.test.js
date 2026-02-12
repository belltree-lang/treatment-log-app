const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const configCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'config.gs'), 'utf8');
const sheetUtilsCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'utils', 'sheetUtils.js'), 'utf8');
const getTodayVisitsCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard', 'api', 'getTodayVisits.js'), 'utf8');
const dashboardHtml = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard.html'), 'utf8');

function createApiContext() {
  const context = {
    console,
    Date,
    JSON,
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        const d = new Date(date);
        if (fmt === 'yyyy-MM-dd') {
          const y = d.getUTCFullYear();
          const m = String(d.getUTCMonth() + 1).padStart(2, '0');
          const day = String(d.getUTCDate()).padStart(2, '0');
          return `${y}-${m}-${day}`;
        }
        if (fmt === 'HH:mm') {
          const h = String(d.getUTCHours()).padStart(2, '0');
          const m = String(d.getUTCMinutes()).padStart(2, '0');
          return `${h}:${m}`;
        }
        return d.toISOString();
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(configCode, context);
  vm.runInContext(sheetUtilsCode, context);
  vm.runInContext(getTodayVisitsCode, context);
  return context;
}

function createElement(tagName) {
  return {
    tagName,
    children: [],
    textContent: '',
    className: '',
    style: {},
    dataset: {},
    disabled: false,
    appendChild(child) {
      this.children.push(child);
      return child;
    },
    addEventListener: () => {},
    setAttribute: () => {},
    removeAttribute: () => {},
    classList: { toggle: () => {}, add: () => {}, remove: () => {} }
  };
}

function flattenText(node) {
  if (!node) return '';
  return [node.textContent || '']
    .concat((node.children || []).map(flattenText))
    .join(' ')
    .trim();
}

function createUiContext() {
  const scriptMatch = dashboardHtml.match(/<script>([\s\S]*?)<\/script>/);
  assert(scriptMatch, 'dashboard script block exists');
  const dashboardScript = scriptMatch[1]
    .replace(/const DASHBOARD_TREATMENT_APP_EXEC_URL =[^\n]*\n/, 'var DASHBOARD_TREATMENT_APP_EXEC_URL = "";\n');

  const elements = {
    visitList: createElement('div')
  };
  elements.visitList.innerHTML = '';

  const documentStub = {
    addEventListener: () => {},
    getElementById: (id) => elements[id] || null,
    createElement: (tagName) => createElement(tagName)
  };

  const context = {
    console,
    URLSearchParams,
    window: { location: { search: '' }, open: () => {} },
    document: documentStub
  };
  vm.createContext(context);
  vm.runInContext(dashboardScript, context);
  return { context, elements };
}

function runVisits(context, now, logs, patientInfo, extraOptions) {
  const options = extraOptions || {};
  return context.getTodayVisits({
    now: new Date(now),
    patientInfo: { patients: patientInfo || {} },
    treatmentLogs: { logs, warnings: [] },
    notes: { notes: {}, warnings: [] },
    visiblePatientIds: options.visiblePatientIds || null
  }).visits;
}

(function testVisiblePatientIdsNullShowsAllVisits() {
  const context = createApiContext();
  const visits = runVisits(
    context,
    '2025-02-13T10:00:00Z',
    [
      { timestamp: new Date('2025-02-13T09:00:00Z'), dateKey: '2025-02-13', patientId: 'P001' },
      { timestamp: new Date('2025-02-10T12:00:00Z'), dateKey: '2025-02-10', patientId: 'P002' }
    ],
    { P001: { name: '患者A' }, P002: { name: '患者B' } }
  );

  assert.strictEqual(visits.length, 2, 'visiblePatientIds=null では全件表示する');
})();

(function testVisiblePatientIdsShowsOnlyTargetPatients() {
  const context = createApiContext();
  const visits = runVisits(
    context,
    '2025-02-13T10:00:00Z',
    [
      { timestamp: new Date('2025-02-13T09:00:00Z'), dateKey: '2025-02-13', patientId: 'P001' },
      { timestamp: new Date('2025-02-13T10:00:00Z'), dateKey: '2025-02-13', patientId: 'P002' },
      { timestamp: new Date('2025-02-10T12:00:00Z'), dateKey: '2025-02-10', patientId: 'P003' }
    ],
    { P001: { name: '患者A' }, P002: { name: '患者B' }, P003: { name: '患者C' } },
    { visiblePatientIds: new Set(['P001']) }
  );

  assert.strictEqual(visits.length, 1, '可視患者のみ表示する');
  assert.strictEqual(visits[0].patientId, 'P001');
})();

(function testVisiblePatientIdsCanResultInZeroVisits() {
  const context = createApiContext();
  const visits = runVisits(
    context,
    '2025-02-13T10:00:00Z',
    [
      { timestamp: new Date('2025-02-13T09:00:00Z'), dateKey: '2025-02-13', patientId: 'P001' },
      { timestamp: new Date('2024-12-20T09:00:00Z'), dateKey: '2024-12-20', patientId: 'P001' }
    ],
    { P001: { name: '患者A' } },
    { visiblePatientIds: new Set() }
  );

  assert.strictEqual(visits.length, 0, '可視患者が空なら表示0件になる');
})();

(function testVisiblePatientFilterKeepsTodayAndLatestPastDayLogic() {
  const context = createApiContext();
  const visits = runVisits(
    context,
    '2025-02-13T10:00:00Z',
    [
      { timestamp: new Date('2025-02-13T09:00:00Z'), dateKey: '2025-02-13', patientId: 'P001' },
      { timestamp: new Date('2025-02-10T12:00:00Z'), dateKey: '2025-02-10', patientId: 'P001' },
      { timestamp: new Date('2025-02-07T12:00:00Z'), dateKey: '2025-02-07', patientId: 'P001' }
    ],
    { P001: { name: '患者A' } },
    { visiblePatientIds: new Set(['P001']) }
  );

  const keys = JSON.parse(JSON.stringify(visits.map(v => v.dateKey)));
  assert.deepStrictEqual(keys, ['2025-02-10', '2025-02-13'], '可視患者フィルタ適用後も今日+最新過去1日のロジックを維持する');
})();

(function testWeekendGapShowsFridayOnMonday() {
  const context = createApiContext();
  const visits = runVisits(
    context,
    '2025-02-10T10:00:00Z',
    [
      { timestamp: new Date('2025-02-10T09:00:00Z'), dateKey: '2025-02-10', patientId: 'P001' },
      { timestamp: new Date('2025-02-07T16:00:00Z'), dateKey: '2025-02-07', patientId: 'P002' },
      { timestamp: new Date('2025-02-06T11:00:00Z'), dateKey: '2025-02-06', patientId: 'P003' }
    ],
    { P001: { name: '患者A' }, P002: { name: '患者B' } }
  );

  assert.strictEqual(visits.length, 2, '月曜表示では今日+直近営業日(金曜)のみ表示する');
  const keys = JSON.parse(JSON.stringify(visits.map(v => v.dateKey)));
  assert.deepStrictEqual(keys, ['2025-02-07', '2025-02-10']);
})();

(function testAfterLeavePicksLatestPastDayOnly() {
  const context = createApiContext();
  const visits = runVisits(
    context,
    '2025-02-13T10:00:00Z',
    [
      { timestamp: new Date('2025-02-13T09:00:00Z'), dateKey: '2025-02-13', patientId: 'P001' },
      { timestamp: new Date('2025-02-10T12:00:00Z'), dateKey: '2025-02-10', patientId: 'P002' },
      { timestamp: new Date('2025-02-07T12:00:00Z'), dateKey: '2025-02-07', patientId: 'P003' }
    ],
    { P001: { name: '患者A' }, P002: { name: '患者B' }, P003: { name: '患者C' } }
  );

  const keys = JSON.parse(JSON.stringify(visits.map(v => v.dateKey)));
  assert.deepStrictEqual(keys, ['2025-02-10', '2025-02-13'], '有給明けでも過去は最新1日だけ');
})();

(function testTodayOnly() {
  const context = createApiContext();
  const visits = runVisits(
    context,
    '2025-02-01T10:00:00Z',
    [
      { timestamp: new Date('2025-02-01T09:30:00Z'), dateKey: '2025-02-01', patientId: 'P001' }
    ],
    { P001: { name: '田中 花子' } }
  );

  assert.strictEqual(visits.length, 1, '今日のみの場合は今日の訪問のみ返す');
  assert.strictEqual(visits[0].dateKey, '2025-02-01');
  assert.strictEqual(visits[0].patientName, '田中 花子', 'patientInfoのnameを引き当てる');
})();

(function testTodayAndPreviousDay() {
  const context = createApiContext();
  const visits = runVisits(
    context,
    '2025-02-01T10:00:00Z',
    [
      { timestamp: new Date('2025-02-01T09:30:00Z'), dateKey: '2025-02-01', patientId: 'P001' },
      { timestamp: new Date('2025-01-31T23:00:00Z'), dateKey: '2025-01-31', patientId: 'P002' },
      { timestamp: new Date('2025-01-30T23:00:00Z'), dateKey: '2025-01-30', patientId: 'P003' }
    ],
    { P001: { name: '田中 花子' }, P002: { name: '山田 太郎' } }
  );

  assert.strictEqual(visits.length, 2, '今日＋前日が存在する場合はその2日を返す');
  const keys = JSON.parse(JSON.stringify(visits.map(v => v.dateKey)));
  assert.deepStrictEqual(keys, ['2025-01-31', '2025-02-01']);
  assert.strictEqual(visits[0].patientName, '山田 太郎', '患者名を表示できる');
})();

(function testRenderVisitsUsesPatientNameWithPatientId() {
  const { context, elements } = createUiContext();
  vm.runInContext(
    `dashboardState.data = { todayVisits: [{ time: '09:30', patientId: 'P001', patientName: '田中 花子', noteStatus: '◎' }] };`,
    context
  );

  context.renderVisits();

  const renderedText = flattenText(elements.visitList);
  assert.ok(renderedText.includes('田中 花子（P001）'), '患者名（患者ID）形式で描画する');
})();

console.log('dashboard today visits tests passed');
