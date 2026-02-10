const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const dashboardHtml = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard.html'),
  'utf8'
);

const scriptMatch = dashboardHtml.match(/<script>([\s\S]*?)<\/script>/);
assert(scriptMatch, 'dashboard script block exists');

const dashboardScript = scriptMatch[1]
  .replace(/const DASHBOARD_TREATMENT_APP_EXEC_URL =[^\n]*\n/, 'var DASHBOARD_TREATMENT_APP_EXEC_URL = "";\n');

function createContext() {
  const openCalls = [];
  const context = {
    console,
    URLSearchParams,
    treatmentAppExecUrl: '',
    window: {
      location: { search: '' },
      open: (...args) => openCalls.push(args)
    },
    document: {
      addEventListener: () => {},
      getElementById: () => null,
      createElement: () => ({
        appendChild: () => {},
        setAttribute: () => {},
        addEventListener: () => {},
        classList: { add: () => {} },
        style: {}
      })
    }
  };
  vm.createContext(context);
  vm.runInContext(dashboardScript, context);
  context.__openCalls = openCalls;
  return context;
}

(function testBuildTreatmentAppLinkFromExecUrl() {
  const context = createContext();
  context.DASHBOARD_TREATMENT_APP_EXEC_URL = 'https://script.google.com/a/macros/belltree1102.com/s/DEPLOY123/exec';

  const link = context.buildTreatmentAppLink_('P-001');
  assert.strictEqual(
    link,
    'https://script.google.com/a/macros/belltree1102.com/s/DEPLOY123/exec?view=record&id=P-001',
    'buildTreatmentAppLink_ returns exec URL with view and patient id'
  );
})();


(function testInvalidGoogleusercontentExecUrlIsRejected() {
  const context = createContext();
  context.DASHBOARD_TREATMENT_APP_EXEC_URL = 'https://script.googleusercontent.com/userCodeAppPanel?createOAuthDialog=true#';

  assert.strictEqual(
    context.resolveTreatmentAppExecUrl_(),
    '',
    'googleusercontent userCodeAppPanel URL is treated as invalid'
  );
})();

(function testMissingExecUrlStopsNavigationAndSetsError() {
  const context = createContext();
  context.DASHBOARD_TREATMENT_APP_EXEC_URL = '';
  context.treatmentAppExecUrl = '';

  assert.strictEqual(context.resolveTreatmentAppExecUrl_(), '', 'resolve returns empty string when missing');
  assert.strictEqual(context.buildTreatmentAppLink_('P-002'), '', 'link build returns empty when missing');

  context.navigateToPatient_('P-002');
  assert.strictEqual(context.__openCalls.length, 0, 'navigate does not open a new tab when URL is missing');
  assert.strictEqual(
    vm.runInContext('dashboardState.error', context),
    '施術録WebアプリURLが未設定のため、患者画面を開けません。',
    'navigate sets a clear UI error message when URL is missing'
  );
})();

console.log('dashboard navigation link tests passed');

// 共有用サマリー
// - 変更点: ダッシュボードの患者遷移ヘルパー（URL解決・リンク構築・遷移）の期待動作を検証する単体テストを追加。
// - 理由: exec URL未設定時に誤遷移せず、設定時は正しい新規タブ遷移URLを生成することを担保するため。
// - 影響範囲: tests 配下のNode実行テストのみ。
