const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const dashboardHtml = fs.readFileSync(path.join(__dirname, '..', 'src', 'dashboard.html'), 'utf8');
const scriptMatch = dashboardHtml.match(/<script>([\s\S]*?)<\/script>/);
assert(scriptMatch, 'dashboard script block exists');

const dashboardScript = scriptMatch[1]
  .replace(/const DASHBOARD_TREATMENT_APP_EXEC_URL =[^\n]*\n/, 'var DASHBOARD_TREATMENT_APP_EXEC_URL = "";\n');

function createElement(tagName) {
  return {
    tagName,
    className: '',
    textContent: '',
    children: [],
    style: {},
    dataset: {},
    appendChild(child) {
      this.children.push(child);
      return child;
    },
    set innerHTML(_value) {
      this.children = [];
    },
    get innerHTML() {
      return '';
    },
    addEventListener() {},
    setAttribute(name, value) {
      this[name] = value;
    },
    removeAttribute() {}
  };
}

function createContext() {
  const criticalPatientList = createElement('div');
  const criticalPatientCount = createElement('div');

  const context = {
    console,
    URLSearchParams,
    treatmentAppExecUrl: 'https://script.google.com/macros/s/DEPLOY123/exec',
    window: {
      location: { search: '' },
      open: () => {}
    },
    document: {
      addEventListener: () => {},
      createElement,
      getElementById: (id) => {
        if (id === 'criticalPatientList') return criticalPatientList;
        if (id === 'criticalPatientCount') return criticalPatientCount;
        return null;
      }
    }
  };

  vm.createContext(context);
  vm.runInContext(dashboardScript, context);
  return { context, criticalPatientList, criticalPatientCount };
}

(function testRenderCriticalPatientsShowsNameAndReasonWithNavigationLink() {
  const { context, criticalPatientList, criticalPatientCount } = createContext();
  vm.runInContext(`dashboardState.data = {
    overview: {
      criticalPatients: {
        count: 2,
        items: [
          { patientId: '001', name: '患者A', reason: '同意期限超過' },
          { patientId: '002', name: '患者B', reason: '報告書遅延' }
        ]
      }
    }
  };`, context);

  context.renderCriticalPatients();

  assert.strictEqual(criticalPatientCount.textContent, '2名');
  assert.strictEqual(criticalPatientList.children.length, 2);
  assert.strictEqual(criticalPatientList.children[0].children[0].textContent, '患者A');
  assert.strictEqual(criticalPatientList.children[0].children[1].textContent, '同意期限超過');
  assert.ok(
    criticalPatientList.children[0].children[0].href.includes('?view=record&id=001'),
    '患者名リンクは既存詳細画面に遷移できる形式'
  );
})();

console.log('dashboard critical patients rendering tests passed');
