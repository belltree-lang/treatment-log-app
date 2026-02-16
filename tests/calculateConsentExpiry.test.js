const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(code, sandbox);

sandbox.Session = {
  getScriptTimeZone: () => 'Asia/Tokyo'
};

sandbox.Utilities = {
  formatDate: date => {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }
};

(function testMonthEndCrossingCases() {
  const cases = [
    { input: '2024-01-10', expected: '2024-06-30' },
    { input: '2024-01-20', expected: '2024-07-31' },
    { input: '2024-03-31', expected: '2024-09-30' },
    { input: '2024-08-16', expected: '2025-02-28' },
    { input: '令和7年6月26日', expected: '2025-12-31' },
    { input: '令和7年1月10日', expected: '2025-06-30' }
  ];

  cases.forEach(({ input, expected }) => {
    assert.strictEqual(sandbox.calculateConsentExpiry_(input), expected, `input=${input}`);
  });
})();

(function testInvalidDateReturnsEmptyString() {
  const result = sandbox.calculateConsentExpiry_('invalid-date');
  assert.strictEqual(result, '');
})();

console.log('calculateConsentExpiry tests passed');
