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
    dataset: {},
    style: {},
    _innerHTML: '',
    appendChild(child) {
      this.children.push(child);
      return child;
    },
    set innerHTML(value) {
      this._innerHTML = value;
      this.children = [];
    },
    get innerHTML() {
      return this._innerHTML;
    },
    addEventListener() {},
    setAttribute() {},
    get childElementCount() {
      return this.children.length;
    }
  };
}

function createContext() {
  const patientList = createElement('div');
  const patientCount = createElement('div');

  const context = {
    console,
    URL,
    URLSearchParams,
    treatmentAppExecUrl: '',
    window: {
      location: { search: '' },
      open: () => {}
    },
    document: {
      addEventListener: () => {},
      createElement,
      getElementById: (id) => {
        if (id === 'patientList') return patientList;
        if (id === 'patientCount') return patientCount;
        return null;
      }
    }
  };

  vm.createContext(context);
  vm.runInContext(dashboardScript, context);
  return { context, patientList };
}

(function testStatusTagsContainerOnlyWhenTagsExist() {
  const { context, patientList } = createContext();
  vm.runInContext(`dashboardState.data = {
    patients: [
      { patientId: '1', name: 'あり', statusTags: [{ type: 'todo', label: '要対応' }] },
      { patientId: '2', name: 'なし', statusTags: [] },
      { patientId: '3', name: '未定義' }
    ]
  };`, context);

  context.renderPatients();

  assert.strictEqual(patientList.children.length, 3, '3 patient rows rendered');

  const headerWithTag = patientList.children[0].children[0].children[0];
  const headerWithoutTag = patientList.children[1].children[0].children[0];
  const headerWithoutTagUndefined = patientList.children[2].children[0].children[0];

  const withTagContainers = headerWithTag.children.filter((child) => child.className === 'status-tags');
  const withoutTagContainers = headerWithoutTag.children.filter((child) => child.className === 'status-tags');
  const withoutTagUndefinedContainers = headerWithoutTagUndefined.children.filter((child) => child.className === 'status-tags');

  assert.strictEqual(withTagContainers.length, 2, 'patient with status tag has status-tag container and info tags container');
  assert.strictEqual(withTagContainers[0].children.length, 1, 'status tag container has one child');

  assert.strictEqual(withoutTagContainers.length, 1, 'patient with empty statusTags only has info tags container');
  assert.strictEqual(withoutTagUndefinedContainers.length, 1, 'patient without statusTags only has info tags container');
})();

console.log('dashboard patient status tags rendering tests passed');
