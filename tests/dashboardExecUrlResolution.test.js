const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const dashboardMainCode = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard', 'main.gs'),
  'utf8'
);

function createContext(options = {}) {
  const props = options.properties || {};
  const context = {
    console,
    DASHBOARD_TREATMENT_APP_EXEC_URL: options.constant || '',
    ScriptApp: {
      getService: () => ({
        getUrl: () => options.serviceUrl || ''
      })
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: key => props[key] || ''
      })
    }
  };
  vm.createContext(context);
  vm.runInContext(dashboardMainCode, context);
  return context;
}

(function testConstantExecUrlWins() {
  const context = createContext({
    constant: 'https://script.google.com/a/macros/belltree1102.com/s/CONST/exec',
    properties: { DASHBOARD_TREATMENT_APP_EXEC_URL: 'https://script.google.com/a/macros/belltree1102.com/s/PROP/exec' },
    serviceUrl: 'https://script.google.com/a/macros/belltree1102.com/s/SVC/exec'
  });
  assert.strictEqual(
    context.resolveDashboardTreatmentAppExecUrl_(),
    'https://script.google.com/a/macros/belltree1102.com/s/CONST/exec'
  );
})();

(function testScriptPropertyUsedWhenConstantMissing() {
  const context = createContext({
    constant: '',
    properties: { DASHBOARD_TREATMENT_APP_EXEC_URL: 'https://script.google.com/a/macros/belltree1102.com/s/PROP/exec' },
    serviceUrl: 'https://script.google.com/a/macros/belltree1102.com/s/SVC/exec'
  });
  assert.strictEqual(
    context.resolveDashboardTreatmentAppExecUrl_(),
    'https://script.google.com/a/macros/belltree1102.com/s/PROP/exec'
  );
})();

(function testGoogleusercontentRejectedAndFallsBackToServiceUrl() {
  const context = createContext({
    constant: 'https://script.googleusercontent.com/userCodeAppPanel?createOAuthDialog=true#',
    serviceUrl: 'https://script.google.com/a/macros/belltree1102.com/s/SVC/exec'
  });
  assert.strictEqual(
    context.resolveDashboardTreatmentAppExecUrl_(),
    'https://script.google.com/a/macros/belltree1102.com/s/SVC/exec'
  );
})();

console.log('dashboard exec URL resolution tests passed');

// 共有用サマリー
// - 変更点: サーバー側のダッシュボードexec URL解決（定数優先・Script Properties・service URLフォールバック）を検証するテストを追加。
// - 理由: userCodeAppPanel系URLの混入を防ぎ、公開exec URLが確実に注入されることを担保するため。
// - 影響範囲: tests 配下のNode実行テストのみ。
