/**
 * AI報告書シートから患者ごとの最終発行日時を取得する。
 * @return {{reports: Object<string, string>, warnings: string[]}}
 */
function loadAIReports(options) {
  const opts = options || {};
  const fetchFn = () => loadAIReportsUncached_();
  if (opts && opts.cache === false) return fetchFn();
  return dashboardCacheFetch_('dashboard:aiReports:v1', fetchFn, DASHBOARD_CACHE_TTL_SECONDS);
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

if (typeof dashboardGetSpreadsheet_ === 'undefined') {
  function dashboardGetSpreadsheet_() {
    if (typeof ss === 'function') {
      try { return ss(); } catch (e) { /* ignore */ }
    }
    if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet) {
      return SpreadsheetApp.getActiveSpreadsheet();
    }
    return null;
  }
}

if (typeof DASHBOARD_CACHE_TTL_SECONDS === 'undefined') {
  var DASHBOARD_CACHE_TTL_SECONDS = 60 * 60 * 12;
}

if (typeof dashboardCacheFetch_ === 'undefined') {
  function dashboardCacheFetch_(cacheKey, fetchFn) {
    if (typeof fetchFn === 'function') return fetchFn();
    return null;
  }
}

if (typeof dashboardParseTimestamp_ === 'undefined') {
  function dashboardParseTimestamp_(value) {
    if (value instanceof Date) return value;
    if (typeof value === 'number' && Number.isFinite(value)) return new Date(value);
    const str = String(value == null ? '' : value).trim();
    if (!str) return null;
    const parsed = new Date(str);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
}

if (typeof dashboardResolveTimeZone_ === 'undefined') {
  function dashboardResolveTimeZone_() {
    if (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function') {
      const tz = Session.getScriptTimeZone();
      if (tz) return tz;
    }
    return 'Asia/Tokyo';
  }
}

if (typeof dashboardFormatDate_ === 'undefined') {
  function dashboardFormatDate_(date, tz, format) {
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
      try { return Utilities.formatDate(date, tz, format); } catch (e) { /* ignore */ }
    }
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
    const iso = date.toISOString();
    if (format && format.indexOf('HH') >= 0) {
      return iso.replace('T', ' ').slice(0, 16);
    }
    return iso.slice(0, 10);
  }
}

if (typeof dashboardNormalizePatientId_ === 'undefined') {
  function dashboardNormalizePatientId_(value) {
    const raw = value == null ? '' : value;
    return String(raw).trim();
  }
}

if (typeof dashboardResolveColumn_ === 'undefined') {
  function dashboardResolveColumn_(headers, candidates, fallbackIndex) {
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
  }
}

if (typeof dashboardWarn_ === 'undefined') {
  function dashboardWarn_(message) {
    if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
      try { Logger.log(message); return; } catch (e) { /* ignore */ }
    }
    if (typeof console !== 'undefined' && console && typeof console.warn === 'function') {
      console.warn(message);
    }
  }
}
