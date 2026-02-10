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
  template.treatmentAppExecUrl = resolveDashboardTreatmentAppExecUrl_();

  return template
    .evaluate()
    .setTitle('施術者ダッシュボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function resolveDashboardTreatmentAppExecUrl_() {
  const configured = typeof DASHBOARD_TREATMENT_APP_EXEC_URL !== 'undefined'
    ? String(DASHBOARD_TREATMENT_APP_EXEC_URL || '').trim()
    : '';
  if (configured) return configured;
  return ScriptApp.getService().getUrl() || '';
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
