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

(function testConsentDateDay1To15AddsFiveMonthsAndRoundsToMonthEnd() {
  const result = sandbox.calculateConsentExpiry_('2025-01-15');
  assert.strictEqual(result, '2025-06-30');
})();

(function testConsentDateDay16OrLaterAddsSixMonthsAndRoundsToMonthEnd() {
  const result = sandbox.calculateConsentExpiry_('2025-01-16');
  assert.strictEqual(result, '2025-07-31');
})();

(function testInvalidDateReturnsEmptyString() {
  const result = sandbox.calculateConsentExpiry_('invalid-date');
  assert.strictEqual(result, '');
})();

console.log('calculateConsentExpiry tests passed');
