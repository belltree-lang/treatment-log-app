/**
 * 申し送りシートを読み込み、患者IDごとの最新申し送りを返す。
 *
 * 旧パス互換のため、utils/sheetUtils.js で提供される一部のヘルパーを
 * 必要に応じて定義する。
 */

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

if (typeof DASHBOARD_CACHE_TTL_SECONDS === 'undefined') {
  var DASHBOARD_CACHE_TTL_SECONDS = 60 * 60 * 12;
}

if (typeof dashboardCoerceDate_ === 'undefined') {
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
}

if (typeof dashboardNormalizePatientId_ === 'undefined') {
  function dashboardNormalizePatientId_(value) {
    const raw = value == null ? '' : value;
    return String(raw).trim();
  }
}

if (typeof dashboardNormalizeEmail_ === 'undefined') {
  function dashboardNormalizeEmail_(email) {
    const raw = email == null ? '' : email;
    const normalized = String(raw).trim().toLowerCase();
    return normalized || '';
  }
}

if (typeof dashboardFormatDate_ === 'undefined') {
  function dashboardFormatDate_(date, tz, format) {
    const targetFormat = format || (typeof DASHBOARD_DATE_FORMAT !== 'undefined' ? DASHBOARD_DATE_FORMAT : 'yyyy-MM-dd');
    const targetTz = tz || (typeof dashboardResolveTimeZone_ === 'function' ? dashboardResolveTimeZone_() : 'Asia/Tokyo');

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

    let result = targetFormat;
    result = result.replace(/yyyy/g, pad(year, 4));
    result = result.replace(/MM/g, month);
    result = result.replace(/dd/g, day);
    result = result.replace(/HH/g, hour);
    result = result.replace(/mm/g, minute);
    result = result.replace(/ss/g, second);
    return result;
  }
}

function loadNotes(options) {
  const opts = options || {};
  const email = opts.email;
  const fetchFn = () => loadNotesUncached_(opts);
  const base = opts && opts.cache === false
    ? fetchFn()
    : dashboardCacheFetch_(dashboardCacheKey_('notes:v1'), fetchFn, DASHBOARD_CACHE_TTL_SECONDS);

  const lastReadAt = loadHandoverLastRead_(email);
  const notes = {};
  Object.keys(base.notes || {}).forEach(pid => {
    const note = base.notes[pid] || {};
    const readAt = lastReadAt[pid] || '';
    const readTs = readAt ? dashboardParseTimestamp_(readAt) : null;
    const noteTs = note.timestamp instanceof Date ? note.timestamp : dashboardParseTimestamp_(note.timestamp || note.when);
    const unread = noteTs ? (!readTs || noteTs.getTime() > readTs.getTime()) : false;
    notes[pid] = Object.assign({}, note, { lastReadAt: readAt, unread });
  });

  const warnings = [];
  if (Array.isArray(base.warnings)) warnings.push.apply(warnings, base.warnings);
  return { notes, warnings, lastReadAt };
}

function loadNotesUncached_(_options) {
  const warnings = [];
  const latestByPatient = {};

  const wb = dashboardGetSpreadsheet_();
  const sheet = wb && wb.getSheetByName ? wb.getSheetByName('申し送り') : null;
  if (!sheet) {
    warnings.push('申し送りシートが見つかりません');
    dashboardWarn_('[loadNotes] sheet not found');
    return { notes: latestByPatient, warnings };
  }

  const lastRow = sheet.getLastRow ? sheet.getLastRow() : 0;
  if (lastRow < 2) return { notes: latestByPatient, warnings };

  const lastCol = Math.max(5, sheet.getLastColumn ? sheet.getLastColumn() : (sheet.getMaxColumns ? sheet.getMaxColumns() : 0));
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const displayValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const tz = typeof dashboardResolveTimeZone_ === 'function' ? dashboardResolveTimeZone_() : 'Asia/Tokyo';

  for (let i = 0; i < values.length; i++) {
    const row = values[i] || [];
    const rowDisplay = displayValues[i] || [];
    const rowNumber = i + 2;

    const patientId = dashboardNormalizePatientId_(row[1] || rowDisplay[1]);
    if (!patientId) {
      warnings.push(`患者IDが未入力の申し送りをスキップしました (row:${rowNumber})`);
      continue;
    }

    const ts = dashboardParseTimestamp_(row[0] || rowDisplay[0]);
    if (!ts) {
      warnings.push(`申し送りの日時を解釈できません (row:${rowNumber})`);
      continue;
    }

    const authorEmail = String(row[2] || rowDisplay[2] || '').trim();
    const body = String(rowDisplay[3] || row[3] || '').trim();
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
        row: rowNumber
      };
    }
  }

  return { notes: latestByPatient, warnings };
}

function dashboardTrimPreview_(text, limit) {
  const raw = String(text == null ? '' : text).replace(/\s+/g, ' ').trim();
  if (!limit || raw.length <= limit) return raw;
  return raw.slice(0, limit);
}

function loadHandoverLastRead_(email) {
  const key = 'HANDOVER_LAST_READ';
  if (typeof PropertiesService === 'undefined' || !PropertiesService.getScriptProperties) return {};
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(key) || '{}';
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return {};
    const emailKey = dashboardNormalizeEmail_(email);
    if (!emailKey) return parsed;
    const result = {};
    Object.keys(parsed).forEach(pid => {
      const entry = parsed[pid];
      if (entry && typeof entry === 'object') {
        if (entry[emailKey]) result[pid] = entry[emailKey];
      } else if (typeof entry === 'string') {
        result[pid] = entry;
      }
    });
    return result;
  } catch (e) {
    dashboardWarn_('[loadHandoverLastRead] parse failed: ' + (e && e.message ? e.message : e));
    return {};
  }
}

function updateHandoverLastRead(patientId, readAt, email) {
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
  const tsIso = ts.toISOString();
  const emailKey = dashboardNormalizeEmail_(email);
  if (emailKey) {
    const existing = current[pid];
    const bucket = existing && typeof existing === 'object' ? existing : {};
    bucket[emailKey] = tsIso;
    current[pid] = bucket;
  } else {
    current[pid] = tsIso;
  }
  try {
    store.setProperty(key, JSON.stringify(current));
    return true;
  } catch (e) {
    dashboardWarn_('[updateHandoverLastRead] failed to save: ' + (e && e.message ? e.message : e));
    return false;
  }
}
