/**
 * 申し送りシートを読み込み、患者IDごとの最新申し送りを返す。
 */
function loadNotes(options) {
  const warnings = [];
  const latestByPatient = {};
  const lastReadAt = loadHandoverLastRead_();
  const opts = options || {};

  const wb = opts.workbook || dashboardGetSpreadsheet_();
  const sheet = wb && wb.getSheetByName ? wb.getSheetByName('申し送り') : null;
  if (!sheet) {
    warnings.push('申し送りシートが見つかりません');
    dashboardWarn_('[loadNotes] sheet not found');
    return { notes: latestByPatient, warnings, lastReadAt };
  }

  const lastRow = sheet.getLastRow ? sheet.getLastRow() : 0;
  if (lastRow < 2) return { notes: latestByPatient, warnings, lastReadAt };

  const lastCol = Math.max(5, sheet.getLastColumn ? sheet.getLastColumn() : (sheet.getMaxColumns ? sheet.getMaxColumns() : 0));
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  const colTimestamp = dashboardResolveColumn_(headers, ['日時', '日付', 'timestamp', 'ts', 'タイムスタンプ'], 1);
  const colPatientId = dashboardResolveColumn_(headers, ['患者ID', 'patientId', 'id'], 2);
  const colAuthor = dashboardResolveColumn_(headers, ['投稿者', 'author', 'authorEmail', 'メール', 'email', '内容'], 3);
  const colBody = dashboardResolveColumn_(headers, ['本文', 'note', '申し送り', 'body', '画像url', '画像URL'], 4);

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const displayValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const tz = dashboardResolveTimeZone_();

  for (let i = 0; i < values.length; i++) {
    const row = values[i] || [];
    const rowDisplay = displayValues[i] || [];
    const rowNumber = i + 2;

    const patientId = dashboardNormalizePatientId_(colPatientId ? (row[colPatientId - 1] || rowDisplay[colPatientId - 1]) : '');
    if (!patientId) {
      warnings.push(`患者IDが未入力の申し送りをスキップしました (row:${rowNumber})`);
      continue;
    }

    const ts = dashboardParseTimestamp_(colTimestamp ? (row[colTimestamp - 1] || rowDisplay[colTimestamp - 1]) : '');
    if (!ts) {
      warnings.push(`申し送りの日時を解釈できません (row:${rowNumber})`);
      continue;
    }

    const authorEmail = String(colAuthor ? (row[colAuthor - 1] || rowDisplay[colAuthor - 1] || '') : '').trim();
    const body = String(colBody ? (rowDisplay[colBody - 1] || row[colBody - 1] || '') : '').trim();
    const preview = dashboardTrimPreview_(body, 20);
    const when = dashboardFormatDate_(ts, tz, 'yyyy-MM-dd HH:mm');

    const existing = latestByPatient[patientId];
    if (!existing || (existing.timestamp && existing.timestamp < ts)) {
      latestByPatient[patientId] = {
        patientId,
        authorEmail,
        note: body,
        preview,
        when,
        timestamp: ts,
        row: rowNumber,
        lastReadAt: lastReadAt[patientId] || ''
      };
    }
  }

  return { notes: latestByPatient, warnings, lastReadAt };
}

function dashboardTrimPreview_(text, limit) {
  const raw = String(text == null ? '' : text).replace(/\s+/g, ' ').trim();
  if (!limit || raw.length <= limit) return raw;
  return raw.slice(0, limit);
}

function loadHandoverLastRead_() {
  const key = 'HANDOVER_LAST_READ';
  if (typeof PropertiesService === 'undefined' || !PropertiesService.getScriptProperties) return {};
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(key) || '{}';
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (e) {
    dashboardWarn_('[loadHandoverLastRead] parse failed: ' + (e && e.message ? e.message : e));
    return {};
  }
}

function updateHandoverLastRead(patientId, readAt) {
  const key = 'HANDOVER_LAST_READ';
  const pid = dashboardNormalizePatientId_(patientId);
  if (!pid) return false;
  if (typeof PropertiesService === 'undefined' || !PropertiesService.getScriptProperties) return false;
  const store = PropertiesService.getScriptProperties();
  let current = {};
  try {
    current = JSON.parse(store.getProperty(key) || '{}');
    if (!current || typeof current !== 'object') current = {};
  } catch (e) {
    current = {};
  }
  const ts = readAt instanceof Date ? readAt : (readAt ? dashboardParseTimestamp_(readAt) : new Date());
  if (!(ts instanceof Date) || Number.isNaN(ts.getTime())) return false;
  current[pid] = ts.toISOString();
  try {
    store.setProperty(key, JSON.stringify(current));
    return true;
  } catch (e) {
    dashboardWarn_('[updateHandoverLastRead] failed to save: ' + (e && e.message ? e.message : e));
    return false;
  }
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

if (typeof dashboardResolveTimeZone_ === 'undefined') {
  function dashboardResolveTimeZone_() {
    if (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function') {
      const tz = Session.getScriptTimeZone();
      if (tz) return tz;
    }
    return 'Asia/Tokyo';
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

if (typeof dashboardFormatDate_ === 'undefined') {
  function dashboardFormatDate_(date, tz, format) {
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
      try { return Utilities.formatDate(date, tz, format); } catch (e) { /* ignore */ }
    }
    return date.toISOString();
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
