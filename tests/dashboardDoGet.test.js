const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const dashboardMainCode = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard', 'main.gs'),
  'utf8'
);

function createContext() {
  const context = { console };
  vm.createContext(context);
  vm.runInContext(dashboardMainCode, context);
  return context;
}

(function testDoGetDelegatesToHandler() {
  const context = createContext();
  assert.strictEqual(typeof context.doGet, 'function', 'doGet is defined');

  const marker = { view: 'dashboard', mock: 'value' };
  let received;
  context.handleDashboardDoGet_ = e => {
    received = e;
    return { delegated: true, marker: e };
  };

  const response = context.doGet(marker);
  assert.strictEqual(received, marker, 'doGet passes the request object to the handler');
  assert.deepStrictEqual(
    response,
    { delegated: true, marker },
    'doGet returns the handler response'
  );
})();

console.log('dashboard doGet delegation test passed');
