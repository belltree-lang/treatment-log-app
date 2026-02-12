const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const html = fs.readFileSync(path.join(__dirname, '../src/app.html'), 'utf8');
const loadNewsMatch = html.match(/function loadNews\(patientId, next\)\{[\s\S]*?\n\}/);
assert.ok(loadNewsMatch, 'loadNews should exist in app.html');

const script = `${loadNewsMatch[0]}\nthis.loadNews = loadNews;`;

async function runInactiveRequestCase() {
  const newsEl = { innerHTML: '' };
  let renderCalls = 0;
  let renderedList = null;
  const context = {
    NEWS_RENDER_TIMEOUT_MS: 100,
    q: id => (id === 'news' ? newsEl : null),
    pid: () => 'P001',
    isActivePatientInfoRequest: () => false,
    renderNewsList: (list) => {
      renderCalls += 1;
      renderedList = list;
    },
    google: {
      script: {
        run: {
          withSuccessHandler(handler) {
            this._success = handler;
            return this;
          },
          withFailureHandler(handler) {
            this._failure = handler;
            return this;
          },
          getNews() {
            setTimeout(() => this._success([{ title: 'new' }]), 0);
          }
        }
      }
    },
    console,
    setTimeout,
    clearTimeout,
    Promise
  };

  vm.createContext(context);
  vm.runInContext(script, context);

  await context.loadNews('P001', { requestToken: 'stale-token' });

  assert.strictEqual(renderCalls, 1, 'news should render even for stale tokens');
  assert.deepStrictEqual(renderedList, [{ title: 'new' }], 'success payload should be rendered');
}

runInactiveRequestCase().then(() => {
  console.log('loadNews inactive request always-render test passed');
}).catch(err => {
  console.error(err);
  process.exit(1);
});
