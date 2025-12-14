/**
 * AI報告書シートから患者ごとの最終発行日時を取得する。
 * @return {{reports: Object<string, string>, warnings: string[]}}
 */
if (typeof dashboardWarn_ !== 'function') {
  var dashboardWarn_ = function(message) {
    if (typeof console !== 'undefined' && console && typeof console.warn === 'function') {
      console.warn(message);
    }
  };
}

if (typeof dashboardResolveColumn_ !== 'function') {
  var dashboardResolveColumn_ = function(headers, candidates, fallbackIndex) {
    if (Array.isArray(candidates)) {
      const normalizedHeaders = (headers || []).map(h => String(h || '').trim().toLowerCase());
      for (let i = 0; i < normalizedHeaders.length; i++) {
        const header = normalizedHeaders[i];
        if (!header) continue;
        if (candidates.some(c => header === String(c || '').trim().toLowerCase())) {
          return i + 1;
        }
      }
    }
    return fallbackIndex || 0;
  };
}

if (typeof dashboardNormalizePatientId_ !== 'function') {
  var dashboardNormalizePatientId_ = function(value) {
    const raw = value == null ? '' : value;
    return String(raw).trim();
  };
}

if (typeof dashboardParseTimestamp_ !== 'function') {
  var dashboardParseTimestamp_ = function(value) {
    if (value instanceof Date) return value;
    if (typeof value === 'number' && Number.isFinite(value)) return new Date(value);
    const str = String(value == null ? '' : value).trim();
    if (!str) return null;
    const parsed = new Date(str);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  };
}

if (typeof dashboardResolveTimeZone_ !== 'function') {
  var dashboardResolveTimeZone_ = function() {
    if (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function') {
      const tz = Session.getScriptTimeZone();
      if (tz) return tz;
    }
    if (typeof DASHBOARD_TIME_ZONE !== 'undefined') return DASHBOARD_TIME_ZONE;
    return 'Asia/Tokyo';
  };
}

if (typeof dashboardFormatDate_ !== 'function') {
  var dashboardFormatDate_ = function(date, tz, format) {
    const targetFormat = format || 'yyyy-MM-dd';
    const targetTz = tz || dashboardResolveTimeZone_();
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
      try { return Utilities.formatDate(date, targetTz, targetFormat); } catch (e) { /* ignore */ }
    }
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
    return date.toISOString();
  };
}

if (typeof DASHBOARD_CACHE_TTL_SECONDS === 'undefined') {
  var DASHBOARD_CACHE_TTL_SECONDS = 300;
}

function loadAIReports(options) {
  const opts = options || {};
  const fetchFn = () => loadAIReportsUncached_();
  if (opts && opts.cache === false) return fetchFn();
  if (typeof dashboardCacheFetch_ !== 'function' || typeof dashboardCacheKey_ !== 'function') {
    return fetchFn();
  }
  return dashboardCacheFetch_(dashboardCacheKey_('aiReports:v1'), fetchFn, DASHBOARD_CACHE_TTL_SECONDS);
}

function loadAIReportsUncached_() {
  const reports = {};
  const warnings = [];

  const wb = dashboardGetSpreadsheet_();
  if (!wb || typeof wb.getSheetByName !== 'function') {
    warnings.push('スプレッドシートを取得できませんでした');
    dashboardWarn_('[loadAIReports] spreadsheet unavailable');
    return { reports, warnings };
  }

  const sheet = wb.getSheetByName('AI報告書');
  if (!sheet) {
    warnings.push('AI報告書シートが見つかりません');
    dashboardWarn_('[loadAIReports] sheet not found');
    return { reports, warnings };
  }

  const lastRow = sheet.getLastRow ? sheet.getLastRow() : 0;
  if (lastRow < 2) return { reports, warnings };

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

  return { reports, warnings };
}
