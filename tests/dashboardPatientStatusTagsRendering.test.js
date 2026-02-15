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
    classList: { add() {}, remove() {}, toggle() {} },
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
      {
        patientId: '1',
        name: 'あり',
        statusTags: [
          { type: 'report', label: '作成済' },
          { type: 'consent', label: '📄 要対応', priority: 'low' },
          { type: 'other', label: '不要' }
        ]
      },
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

  assert.strictEqual(withTagContainers.length, 1, 'patient with status tag has one status-tags container');
  assert.strictEqual(withTagContainers[0].children.length, 2, 'only consent/report tags are rendered');
  assert.strictEqual(withTagContainers[0].children[0].textContent, '📄 要対応', 'consent is displayed before report');
  assert.strictEqual(withTagContainers[0].children[1].textContent, '作成済', 'report is displayed after consent');

  assert.strictEqual(withoutTagContainers.length, 0, 'patient with empty statusTags has no status-tags container');
  assert.strictEqual(withoutTagUndefinedContainers.length, 0, 'patient without statusTags has no status-tags container');
})();

(function testReportTagUsesPendingDoneClass() {
  const { context, patientList } = createContext();
  vm.runInContext(`dashboardState.data = {
    patients: [
      { patientId: '1', name: '未作成', statusTags: [{ type: 'report', label: '未作成' }] },
      { patientId: '2', name: '作成済', statusTags: [{ type: 'report', label: '作成済' }] }
    ]
  };`, context);

  context.renderPatients();

  const allTags = patientList.children.map((row) => {
    const header = row.children[0].children[0];
    return header.children.find((child) => child.className === 'status-tags').children[0];
  });
  const pendingTag = allTags.find((tag) => tag.textContent === '未作成');
  const doneTag = allTags.find((tag) => tag.textContent === '作成済');

  assert.ok(pendingTag.className.includes('tag-report-pending'), '未作成 should use pending class');
  assert.ok(doneTag.className.includes('tag-report-done'), '作成済 should use done class');
})();


(function testConsentPriorityUsesExpectedClassNames() {
  const { context, patientList } = createContext();
  vm.runInContext(`dashboardState.data = {
    patients: [
      { patientId: '1', name: '高', statusTags: [{ type: 'consent', label: '⚠ 期限超過', priority: 'high' }] },
      { patientId: '2', name: '中', statusTags: [{ type: 'consent', label: '⏳ 期限迫る', priority: 'medium' }] },
      { patientId: '3', name: '低', statusTags: [{ type: 'consent', label: '📄 要対応', priority: 'low' }] }
    ]
  };`, context);

  context.renderPatients();

  const tags = patientList.children.map((row) => {
    const header = row.children[0].children[0];
    return header.children.find((child) => child.className === 'status-tags').children[0];
  });

  const highTag = tags.find((tag) => tag.textContent === '⚠ 期限超過');
  const mediumTag = tags.find((tag) => tag.textContent === '⏳ 期限迫る');
  const lowTag = tags.find((tag) => tag.textContent === '📄 要対応');

  assert.ok(highTag.className.includes('status-consent-expired'), 'high priority should map to status-consent-expired');
  assert.ok(mediumTag.className.includes('status-consent-warning'), 'medium priority should map to status-consent-warning');
  assert.ok(lowTag.className.includes('status-consent-normal'), 'low priority should map to status-consent-normal');
})();

(function testOverviewListDoesNotRenderCountText() {
  const { context } = createContext();
  const list = context.renderOverviewList_([
    { patientId: '1', name: '患者A', count: 3, subText: '補足' }
  ], '対象なし');

  const title = list.children[0].children[0];
  const titleTexts = title.children.map((child) => child.textContent);

  assert.deepStrictEqual(titleTexts, ['患者A'], '患者名のみ表示し件数は表示しない');
})();

console.log('dashboard patient status tags rendering tests passed');
