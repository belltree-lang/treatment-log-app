/**
 * AI報告書シートから患者ごとの最終発行日時を取得する。
 * @return {{reports: Object<string, string>, warnings: string[]}}
 */
function loadAIReports(options) {
  const opts = options || {};
  const fetchFn = () => loadAIReportsUncached_();
  return dashboardCacheFetch_(dashboardCacheKey_('aiReports:v1'), fetchFn, DASHBOARD_CACHE_TTL_SECONDS, opts);
}

function loadAIReportsUncached_() {
  const reports = {};
  const warnings = [];
  let setupIncomplete = false;

  const wb = dashboardGetSpreadsheet_();
  if (!wb || typeof wb.getSheetByName !== 'function') {
    warnings.push('スプレッドシートを取得できませんでした');
    setupIncomplete = true;
    dashboardWarn_('[loadAIReports] spreadsheet unavailable');
    return { reports, warnings, setupIncomplete };
  }

  const sheetName = typeof DASHBOARD_SHEET_AI_REPORTS !== 'undefined' ? DASHBOARD_SHEET_AI_REPORTS : 'AI報告書';
  const sheet = wb.getSheetByName(sheetName);
  if (!sheet) {
    const warning = `${sheetName}シートが見つかりません`;
    warnings.push(warning);
    setupIncomplete = true;
    dashboardWarn_(`[loadAIReports] sheet not found: ${sheetName}`);
    return { reports, warnings, setupIncomplete };
  }

  const lastRow = sheet.getLastRow ? sheet.getLastRow() : 0;
  if (lastRow < 2) return { reports, warnings, setupIncomplete };

  const lastCol = Math.max(2, sheet.getLastColumn ? sheet.getLastColumn() : (sheet.getMaxColumns ? sheet.getMaxColumns() : 2));
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  const tsCol = dashboardResolveColumn_(headers, ['ts', 'timestamp', '日時'], 1);
  const patientCol = dashboardResolveColumn_(headers, ['患者id', '患者ＩＤ', '患者ID', 'patientId', 'id'], 2);

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const displays = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const tz = dashboardResolveTimeZone_();
  const latestTs = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i] || [];
    const disp = displays[i] || [];

    const patientId = dashboardNormalizePatientId_(disp[patientCol - 1] || row[patientCol - 1]);
    if (!patientId) continue;

    const tsValue = row[tsCol - 1] != null && row[tsCol - 1] !== '' ? row[tsCol - 1] : disp[tsCol - 1];
    const tsDate = dashboardParseTimestamp_(tsValue);
    if (!tsDate) continue;

    const ts = tsDate.getTime();
    const currentTs = latestTs[patientId] || 0;
    if (ts <= currentTs) continue;

    latestTs[patientId] = ts;
    reports[patientId] = dashboardFormatDate_(tsDate, tz, 'yyyy-MM-dd HH:mm');
  }

  return { reports, warnings, setupIncomplete };
}
