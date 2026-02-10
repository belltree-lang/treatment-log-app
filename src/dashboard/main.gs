function handleDashboardDoGet_(e) {
  if (!shouldHandleDashboardRequest_(e)) return null;

  if (shouldHandleDashboardApi_(e)) {
    const user = dashboardResolveRequestUser_(e);
    const useMock = shouldUseDashboardMockData_(e);
    const data = useMock && typeof getDashboardMockData === 'function'
      ? getDashboardMockData({ user })
      : typeof getDashboardData === 'function'
        ? getDashboardData({ user })
        : {};
    return createJsonResponse_(data);
  }

  const template = HtmlService.createTemplateFromFile('dashboard');
  template.baseUrl = ScriptApp.getService().getUrl() || '';
  template.dashboardTreatmentAppExecUrlConstant = resolveDashboardTreatmentAppExecUrlFromConstant_();
  template.treatmentAppExecUrl = resolveDashboardTreatmentAppExecUrl_();

  return template
    .evaluate()
    .setTitle('施術者ダッシュボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function resolveDashboardTreatmentAppExecUrl_() {
  const configured = resolveDashboardTreatmentAppExecUrlFromConstant_();
  if (configured) return configured;

  const serviceUrl = normalizeDashboardTreatmentAppExecUrl_(ScriptApp.getService().getUrl() || '');
  if (serviceUrl) return serviceUrl;

  return '';
}

function resolveDashboardTreatmentAppExecUrlFromConstant_() {
  if (typeof DASHBOARD_TREATMENT_APP_EXEC_URL !== 'undefined') {
    const fromConstant = normalizeDashboardTreatmentAppExecUrl_(String(DASHBOARD_TREATMENT_APP_EXEC_URL || '').trim());
    if (fromConstant) return fromConstant;
  }

  return resolveDashboardTreatmentAppExecUrlFromProperties_();
}

function resolveDashboardTreatmentAppExecUrlFromProperties_() {
  try {
    if (typeof PropertiesService === 'undefined' || !PropertiesService || typeof PropertiesService.getScriptProperties !== 'function') {
      return '';
    }
    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties || typeof scriptProperties.getProperty !== 'function') return '';
    const raw = scriptProperties.getProperty('DASHBOARD_TREATMENT_APP_EXEC_URL') || '';
    return normalizeDashboardTreatmentAppExecUrl_(String(raw || '').trim());
  } catch (err) {
    return '';
  }
}

function normalizeDashboardTreatmentAppExecUrl_(url) {
  const value = String(url || '').trim();
  if (!value) return '';
  const normalized = value.toLowerCase();
  if (normalized.indexOf('googleusercontent.com/usercodeapppanel') >= 0) return '';
  if (normalized.indexOf('script.google.com/') < 0) return '';
  if (normalized.indexOf('/exec') < 0) return '';
  return value;
}

function shouldHandleDashboardRequest_(e) {
  const path = (e && e.pathInfo ? String(e.pathInfo) : '').replace(/^\/+|\/+$/g, '').toLowerCase();
  if (path === 'dashboard' || path === 'getdashboarddata') return true;
  const view = e && e.parameter ? String(e.parameter.view || '').toLowerCase() : '';
  if (view === 'dashboard') return true;
  const action = e && e.parameter ? String(e.parameter.action || e.parameter.api || '').toLowerCase() : '';
  return action === 'getdashboarddata';
}

function shouldHandleDashboardApi_(e) {
  const path = (e && e.pathInfo ? String(e.pathInfo) : '').replace(/^\/+|\/+$/g, '').toLowerCase();
  if (path === 'getdashboarddata') return true;
  const action = e && e.parameter ? (e.parameter.action || e.parameter.api) : '';
  return String(action || '').toLowerCase() === 'getdashboarddata';
}

function dashboardResolveRequestUser_(e) {
  const paramUser = e && e.parameter ? (e.parameter.user || e.parameter.email) : '';
  if (paramUser) return String(paramUser || '').trim();
  if (typeof dashboardResolveUser_ === 'function') return dashboardResolveUser_();
  return '';
}

function shouldUseDashboardMockData_(e) {
  const raw = e && e.parameter ? e.parameter.mock : '';
  const normalized = String(raw || '').trim().toLowerCase();
  return normalized === '1' || normalized === 'true' || normalized === 'yes';
}

function createJsonResponse_(payload) {
  if (typeof ContentService === 'undefined' || !ContentService || typeof ContentService.createTextOutput !== 'function') {
    return JSON.stringify(payload || {});
  }
  const output = ContentService.createTextOutput(JSON.stringify(payload || {}));
  if (output && typeof output.setMimeType === 'function' && ContentService && ContentService.MimeType) {
    output.setMimeType(ContentService.MimeType.JSON);
  }
  return output;
}

// 共有用サマリー
// - 変更点: ダッシュボード向け施術録WebアプリURLの解決元を定数→スクリプトプロパティ→現在デプロイURLに整理し、googleusercontent系URLを除外する正規化を追加。
// - 理由: 患者遷移先が userCodeAppPanel に化ける不具合を防ぎ、常に公開exec URLを優先するため。
// - 影響範囲: dashboard doGetテンプレート注入時のURL解決ロジックのみ。
