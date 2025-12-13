/**
 * 施術録を読み込み、患者名を患者IDへマッピングした結果を返す。
 */
function loadTreatmentLogs(options) {
  const opts = options || {};
  const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo() : null);
  const nameToId = opts.nameToId || (patientInfo && patientInfo.nameToId) || {};
  const warnings = patientInfo && Array.isArray(patientInfo.warnings) ? [].concat(patientInfo.warnings) : [];
  const logs = [];
  const lastStaffByPatient = {};

  const wb = dashboardGetSpreadsheet_();
  const sheet = wb && wb.getSheetByName ? wb.getSheetByName('施術録') : null;
  if (!sheet) {
    warnings.push('施術録シートが見つかりません');
    dashboardWarn_('[loadTreatmentLogs] sheet not found');
    return { logs, warnings, lastStaffByPatient };
  }

  const lastRow = sheet.getLastRow ? sheet.getLastRow() : 0;
  if (lastRow < 2) return { logs, warnings, lastStaffByPatient };

  const lastCol = sheet.getLastColumn ? sheet.getLastColumn() : sheet.getMaxColumns ? sheet.getMaxColumns() : 0;
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const displayValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  const colTimestamp = dashboardResolveColumn_(headers, ['日時', '日付', 'timestamp', 'ts', 'タイムスタンプ'], 1);
  const colPatientId = dashboardResolveColumn_(headers, ['患者ID', 'patientId', 'id'], 0);
  const colPatientName = dashboardResolveColumn_(headers, ['氏名', '名前', '患者名'], colPatientId ? 0 : 2);
  const colEmail = dashboardResolveColumn_(headers, ['作成者', '担当者', 'email', 'createdbyemail'], 0);

  const tz = dashboardResolveTimeZone_();
  const monthStart = dashboardStartOfMonth_(tz, new Date());
  const prevMonthEnd = dashboardEndOfPreviousMonth_(monthStart);

  for (let i = 0; i < values.length; i++) {
    const row = values[i] || [];
    const rowDisplay = displayValues[i] || [];
    const rowNumber = i + 2;

    const timestamp = dashboardParseTimestamp_(row[colTimestamp - 1] || rowDisplay[colTimestamp - 1]);
    if (!timestamp) {
      warnings.push(`施術日時を解釈できません (row:${rowNumber})`);
      continue;
    }

    const patientIdRaw = colPatientId ? dashboardNormalizePatientId_(row[colPatientId - 1] || rowDisplay[colPatientId - 1]) : '';
    const patientName = String((colPatientName && (rowDisplay[colPatientName - 1] || row[colPatientName - 1])) || '').trim();
    const mappedPatientId = patientIdRaw || dashboardResolvePatientIdFromName_(patientName, nameToId);
    const createdByEmail = colEmail ? String(rowDisplay[colEmail - 1] || row[colEmail - 1] || '').trim() : '';
    const dateKey = dashboardFormatDate_(timestamp, tz, 'yyyy-MM-dd');

    if (!mappedPatientId) {
      warnings.push(`患者名をIDに紐付けできません: ${patientName || '(空白)'} (row:${rowNumber})`);
    }

    const entry = {
      row: rowNumber,
      patientId: mappedPatientId || '',
      patientName,
      createdByEmail,
      timestamp,
      dateKey
    };
    if (!mappedPatientId) {
      entry.unmapped = true;
      entry.unmappedName = patientName || '';
    }

    logs.push(entry);

    if (mappedPatientId && timestamp <= prevMonthEnd) {
      const prev = lastStaffByPatient[mappedPatientId];
      if (!prev || (prev.timestamp && prev.timestamp < timestamp)) {
        lastStaffByPatient[mappedPatientId] = { email: createdByEmail, timestamp };
      }
    }
  }

  Object.keys(lastStaffByPatient).forEach(pid => {
    const entry = lastStaffByPatient[pid];
    lastStaffByPatient[pid] = entry ? entry.email || '' : '';
  });

  return { logs, warnings, lastStaffByPatient };
}

function dashboardResolvePatientIdFromName_(name, nameToId) {
  const key = dashboardNormalizeNameKey_(name);
  return key && nameToId ? nameToId[key] : '';
}

function dashboardParseTimestamp_(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'number' && Number.isFinite(value)) return new Date(value);
  const str = String(value == null ? '' : value).trim();
  if (!str) return null;
  const parsed = new Date(str);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function dashboardResolveTimeZone_() {
  if (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function') {
    const tz = Session.getScriptTimeZone();
    if (tz) return tz;
  }
  return 'Asia/Tokyo';
}

function dashboardFormatDate_(date, tz, format) {
  if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
    try { return Utilities.formatDate(date, tz, format); } catch (e) { /* ignore */ }
  }
  return date.toISOString().slice(0, 10);
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

if (typeof dashboardNormalizeNameKey_ === 'undefined') {
  function dashboardNormalizeNameKey_(name) {
    return String(name == null ? '' : name)
      .replace(/\s+/g, '')
      .toLowerCase();
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
