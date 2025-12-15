/**
 * 申し送りシートを読み込み、患者IDごとの最新申し送りを返す。
 */
function loadNotes(options) {
  const opts = options || {};
  const email = opts.email;
  const fetchFn = () => loadNotesUncached_(opts);
  const base = dashboardCacheFetch_(dashboardCacheKey_('notes:v1'), fetchFn, DASHBOARD_CACHE_TTL_SECONDS, opts);

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
  const tz = dashboardResolveTimeZone_();

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
