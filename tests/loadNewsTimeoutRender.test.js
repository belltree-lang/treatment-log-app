// Manual verification (local HTML):
// 1. Open src/app.html in Apps Script preview / local host where google.script.run is available.
// 2. Trigger a patient switch rapidly so older getNews responses become token-mismatched.
// 3. Confirm #news does not stay on "読み込み中…" after success/failure/timeout.

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const html = fs.readFileSync(path.join(__dirname, '../src/app.html'), 'utf8');
const loadNewsMatch = html.match(/function loadNews\(patientId, next\)\{[\s\S]*?\n\}/);
assert.ok(loadNewsMatch, 'loadNews should exist in app.html');

const script = `${loadNewsMatch[0]}\nthis.loadNews = loadNews;`;

async function runTimeoutCase() {
  const newsEl = { innerHTML: '' };
  let renderCalls = 0;
  const context = {
    NEWS_RENDER_TIMEOUT_MS: 5,
    q: id => (id === 'news' ? newsEl : null),
    pid: () => 'P001',
    isActivePatientInfoRequest: () => true,
    renderNewsList: (list) => {
      renderCalls += 1;
      newsEl.innerHTML = Array.isArray(list) && list.length
        ? '<div>rendered</div>'
        : '<div class="muted">Newsはありません</div>';
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
            // Intentionally never resolves to trigger timeout path.
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

  await context.loadNews('P001', { requestToken: 'token-1' });

  assert.ok(!newsEl.innerHTML.includes('読み込み中…'), 'timeout should clear loading message');
  assert.strictEqual(renderCalls, 1, 'timeout should render exactly once');
}

runTimeoutCase().then(() => {
  console.log('loadNews timeout render test passed');
}).catch(err => {
  console.error(err);
  process.exit(1);
});
