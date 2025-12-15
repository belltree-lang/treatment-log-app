/**
 * ダッシュボード共通のユーティリティ群。
 */
function dashboardWarn_(message) {
  if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
    try { Logger.log(message); return; } catch (e) { /* ignore */ }
  }
  if (typeof console !== 'undefined' && console && typeof console.warn === 'function') {
    console.warn(message);
  }
}

function dashboardGetSpreadsheet_() {
  if (typeof DASHBOARD_SPREADSHEET_ID !== 'undefined' && DASHBOARD_SPREADSHEET_ID) {
    try {
      if (typeof SpreadsheetApp !== 'undefined'
        && SpreadsheetApp
        && typeof SpreadsheetApp.openById === 'function') {
        return SpreadsheetApp.openById(DASHBOARD_SPREADSHEET_ID);
      }
    } catch (e) {
      dashboardWarn_('[dashboardGetSpreadsheet] failed to open by ID: ' + (e && e.message ? e.message : e));
    }
  }

  if (typeof ss === 'function') {
    try { return ss(); } catch (e) { /* ignore */ }
  }
  if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  return null;
}

function dashboardGetInvoiceRootFolder_() {
  if (typeof DASHBOARD_INVOICE_FOLDER_ID !== 'undefined' && DASHBOARD_INVOICE_FOLDER_ID) {
    try {
      if (typeof DriveApp !== 'undefined' && DriveApp && typeof DriveApp.getFolderById === 'function') {
        return DriveApp.getFolderById(DASHBOARD_INVOICE_FOLDER_ID);
      }
    } catch (e) {
      dashboardWarn_('[dashboardGetInvoiceRootFolder] failed to open by ID: ' + (e && e.message ? e.message : e));
    }
  }

  if (typeof INVOICE_PARENT_FOLDER_ID !== 'undefined' && INVOICE_PARENT_FOLDER_ID) {
    try {
      if (typeof DriveApp !== 'undefined' && DriveApp && typeof DriveApp.getFolderById === 'function') {
        return DriveApp.getFolderById(INVOICE_PARENT_FOLDER_ID);
      }
    } catch (e) {
      dashboardWarn_('[dashboardGetInvoiceRootFolder] failed to open fallback ID: ' + (e && e.message ? e.message : e));
    }
  }

  return null;
}

function dashboardResolveTimeZone_() {
  if (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function') {
    const tz = Session.getScriptTimeZone();
    if (tz) return tz;
  }
  if (typeof DEFAULT_TZ !== 'undefined') return DEFAULT_TZ;
  return 'Asia/Tokyo';
}

function dashboardFormatDate_(date, tz, format) {
  const targetFormat = format || (typeof DATE_FORMAT !== 'undefined' ? DATE_FORMAT : 'yyyy/MM/dd');
  const targetTz = tz || dashboardResolveTimeZone_();

  if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
    try { return Utilities.formatDate(date, targetTz, targetFormat); } catch (e) { /* ignore */ }
  }

  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  const pad = (value, length) => String(value).padStart(length, '0');
  const year = date.getFullYear();
  const month = pad(date.getMonth() + 1, 2);
  const day = pad(date.getDate(), 2);
  const hour = pad(date.getHours(), 2);
  const minute = pad(date.getMinutes(), 2);
  const second = pad(date.getSeconds(), 2);
  const offsetMinutes = date.getTimezoneOffset();
  const sign = offsetMinutes > 0 ? '-' : '+';
  const absOffset = Math.abs(offsetMinutes);
  const offsetHour = pad(Math.floor(absOffset / 60), 2);
  const offsetMinute = pad(absOffset % 60, 2);

  let result = targetFormat;
  result = result.replace(/yyyy/g, pad(year, 4));
  result = result.replace(/MM/g, month);
  result = result.replace(/dd/g, day);
  result = result.replace(/HH/g, hour);
  result = result.replace(/mm/g, minute);
  result = result.replace(/ss/g, second);
  result = result.replace(/XXX/g, `${sign}${offsetHour}:${offsetMinute}`);
  return result;
}

function dashboardParseTimestamp_(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'number' && Number.isFinite(value)) return new Date(value);
  const str = String(value == null ? '' : value).trim();
  if (!str) return null;
  const parsed = new Date(str);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function dashboardCoerceDate_(value) {
  if (value instanceof Date) return value;
  if (value && typeof value.getTime === 'function') {
    const ts = value.getTime();
    if (Number.isFinite(ts)) return new Date(ts);
  }
  if (value === null || value === undefined) return null;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function dashboardNormalizePatientId_(value) {
  const raw = value == null ? '' : value;
  return String(raw).trim();
}

function dashboardNormalizeNameKey_(name) {
  return String(name == null ? '' : name)
    .replace(/\s+/g, '')
    .toLowerCase();
}

function dashboardNormalizeEmail_(email) {
  const raw = email == null ? '' : email;
  const normalized = String(raw).trim().toLowerCase();
  return normalized || '';
}

function dashboardResolvePatientIdFromName_(name, nameToId) {
  const key = dashboardNormalizeNameKey_(name);
  return key && nameToId ? nameToId[key] : '';
}

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

function dashboardStartOfMonth_(tz, now) {
  const ref = now instanceof Date ? now : new Date();
  const y = ref.getFullYear();
  const m = ref.getMonth();
  return new Date(y, m, 1, 0, 0, 0, 0);
}

function dashboardEndOfPreviousMonth_(monthStart) {
  const base = monthStart instanceof Date ? monthStart : new Date();
  return new Date(base.getFullYear(), base.getMonth(), 0, 23, 59, 59, 999);
}
