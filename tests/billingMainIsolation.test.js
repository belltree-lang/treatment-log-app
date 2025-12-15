const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const mainCode = fs.readFileSync(path.join(__dirname, '..', 'src', 'main.gs'), 'utf8');

function createContext(overrides = {}) {
  const warnings = [];
  let templateState = null;
  const htmlOutput = {
    title: '',
    meta: [],
    frameMode: '',
    setTitle(value) { this.title = value; return this; },
    addMetaTag(name, content) { this.meta.push({ name, content }); return this; },
    setXFrameOptionsMode(mode) { this.frameMode = mode; return this; }
  };

  const context = Object.assign({
    console: { warn: msg => warnings.push(msg) },
    HtmlService: {
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
      createTemplateFromFile: name => {
        templateState = {
          file: name,
          baseUrl: '',
          patientId: '',
          payrollPdfData: null,
          lead: undefined,
          evaluate: () => htmlOutput
        };
        return templateState;
      }
    },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://example.com/app' }) }
  }, overrides);

  vm.createContext(context);
  vm.runInContext(mainCode, context);
  return { context, warnings, htmlOutput, getTemplateState: () => templateState };
}

(function testDoGetUsesOnlyBillingTemplate() {
  const { context, warnings, htmlOutput, getTemplateState } = createContext();
  const response = context.doGet({ pathInfo: 'dashboard', parameter: { view: 'dashboard', id: '010', lead: 'abc' } });

  assert.strictEqual(response, htmlOutput, 'doGet returns the evaluated HtmlOutput');
  const template = getTemplateState();
  assert.strictEqual(template.file, 'billing', 'billing template is always used');
  assert.strictEqual(template.patientId, '010', 'patientId is passed to the template');
  assert.strictEqual(template.lead, 'abc', 'lead parameter is preserved');
  assert.strictEqual(htmlOutput.title, '請求処理アプリ', 'title is fixed to billing app');
  assert.strictEqual(htmlOutput.frameMode, 'ALLOWALL', 'X-Frame-Options mode is configured');
  assert.ok(warnings.length > 0, 'dashboard routes are warned but not executed');
})();

(function testDashboardHooksAreAbsent() {
  const { context } = createContext();
  ['handleDashboardDoGet_', 'shouldHandleDashboardRequest_', 'getDashboardData'].forEach(name => {
    assert.strictEqual(typeof context[name], 'undefined', `${name} should not exist in billing project`);
  });
})();

console.log('billing main isolation tests passed');
