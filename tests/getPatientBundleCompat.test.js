const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(code, sandbox);

sandbox.normId_ = v => String(v || '').trim();
sandbox.getPatientHeader = pid => ({ patientId: pid });
sandbox.getNews = () => [];
sandbox.convertPlainTextToSafeHtml_ = s => s;

const now = new Date();
const expectedYear = now.getFullYear();
const expectedMonth = now.getMonth() + 1;
let called = null;
sandbox.listTreatmentsForMonth = (pid, year, month) => {
  called = { pid, year, month };
  return [];
};

sandbox.getPatientBundle('P001');

assert.ok(called, 'listTreatmentsForMonth should be called');
assert.strictEqual(called.pid, 'P001');
assert.strictEqual(called.year, expectedYear, 'month未指定時は当年を使う');
assert.strictEqual(called.month, expectedMonth, 'month未指定時は当月を使う');

console.log('getPatientBundle compat tests passed');
