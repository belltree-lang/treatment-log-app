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

  const elements = { visitList: createElement('div') };
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

function runVisits(context, now, logs, patientInfo) {
  return context.getTodayVisits({
    now: new Date(now),
    patientInfo: { patients: patientInfo || {} },
    treatmentLogs: { logs, warnings: [] },
    notes: { notes: {}, warnings: [] }
  });
}

(function testTodayAndPreviousBothExist() {
  const context = createApiContext();
  const result = runVisits(
    context,
    '2025-02-13T10:00:00Z',
    [
      { timestamp: new Date('2025-02-13T09:00:00Z'), dateKey: '2025-02-13', patientId: 'P001' },
      { timestamp: new Date('2025-02-12T12:00:00Z'), dateKey: '2025-02-12', patientId: 'P002' },
      { timestamp: new Date('2025-02-11T12:00:00Z'), dateKey: '2025-02-11', patientId: 'P003' }
    ],
    { P001: { name: '患者A' }, P002: { name: '患者B' }, P003: { name: '患者C' } }
  );

  assert.strictEqual(result.today.date, '2025-02-13');
  assert.strictEqual(result.today.visits.length, 1);
  assert.strictEqual(result.previous.date, '2025-02-12');
  assert.strictEqual(result.previous.visits.length, 1);
})();

(function testNoTodayButPreviousExists() {
  const context = createApiContext();
  const result = runVisits(
    context,
    '2025-02-13T10:00:00Z',
    [
      { timestamp: new Date('2025-02-12T09:00:00Z'), dateKey: '2025-02-12', patientId: 'P001' },
      { timestamp: new Date('2025-02-10T09:00:00Z'), dateKey: '2025-02-10', patientId: 'P002' }
    ],
    { P001: { name: '患者A' }, P002: { name: '患者B' } }
  );

  assert.strictEqual(result.today.visits.length, 0);
  assert.strictEqual(result.previous.date, '2025-02-12');
  assert.strictEqual(result.previous.visits.length, 1);
})();

(function testTodayExistsButNoPrevious() {
  const context = createApiContext();
  const result = runVisits(
    context,
    '2025-02-13T10:00:00Z',
    [{ timestamp: new Date('2025-02-13T09:00:00Z'), dateKey: '2025-02-13', patientId: 'P001' }],
    { P001: { name: '患者A' } }
  );

  assert.strictEqual(result.today.visits.length, 1);
  assert.strictEqual(result.previous.date, null);
  assert.strictEqual(result.previous.visits.length, 0);
})();

(function testBothTodayAndPreviousAreEmpty() {
  const context = createApiContext();
  const result = runVisits(context, '2025-02-13T10:00:00Z', [], {});

  assert.strictEqual(result.today.date, '2025-02-13');
  assert.strictEqual(result.today.visits.length, 0);
  assert.strictEqual(result.previous.date, null);
  assert.strictEqual(result.previous.visits.length, 0);
})();

(function testTodayVisitsFollowsProvidedLogsOnly() {
  const context = createApiContext();
  const staffLike = runVisits(
    context,
    '2025-02-13T10:00:00Z',
    [
      { timestamp: new Date('2025-02-13T09:00:00Z'), dateKey: '2025-02-13', patientId: 'P001' },
      { timestamp: new Date('2025-02-12T09:00:00Z'), dateKey: '2025-02-12', patientId: 'P003' }
    ],
    { P001: { name: '患者A' }, P003: { name: '患者C' } }
  );

  assert.strictEqual(staffLike.today.visits.length, 1);
  assert.strictEqual(staffLike.today.visits[0].patientId, 'P001');
  assert.strictEqual(staffLike.previous.visits.length, 1);
  assert.strictEqual(staffLike.previous.visits[0].patientId, 'P003');
})();

(function testRenderVisitsAsTodayAndPreviousSections() {
  const { context, elements } = createUiContext();
  vm.runInContext(
    `dashboardState.data = { todayVisits: { today: { date: '2025-02-13', visits: [{ time: '09:30', patientId: 'P001', patientName: '田中 花子', noteStatus: '◎' }] }, previous: { date: null, visits: [] } } };`,
    context
  );

  context.renderVisits();

  const renderedText = flattenText(elements.visitList);
  assert.ok(renderedText.includes('当日（02/13）'), '当日見出しは MM/DD 形式で描画する');
  assert.ok(renderedText.includes('前日（-）'), '前日の日付 null は - で描画する');
  assert.ok(renderedText.includes('田中 花子（P001）'), '患者名（患者ID）形式で描画する');
  assert.ok(renderedText.includes('0件'), '前日0件を表示する');
})();

console.log('dashboard today visits tests passed');
