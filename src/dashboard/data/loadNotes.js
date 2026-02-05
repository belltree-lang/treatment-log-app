/**
 * 申し送りシートを読み込み、患者IDごとの最新申し送りを返す。
 */
function loadNotes(options) {
  const opts = options || {};
  const email = opts.email;
  const base = loadNotesUncached_(opts);
  const logContext = (label, details) => {
    if (typeof dashboardLogContext_ === 'function') {
      dashboardLogContext_(label, details);
    } else if (typeof dashboardWarn_ === 'function') {
      const payload = details ? ` ${details}` : '';
      dashboardWarn_(`[${label}]${payload}`);
    }
  };

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
  logContext('loadNotes:done', `notes=${Object.keys(notes).length} warnings=${warnings.length} setupIncomplete=${!!base.setupIncomplete}`);
  return { notes, warnings, lastReadAt, setupIncomplete: !!base.setupIncomplete };
}

function loadNotesUncached_(_options) {
  const warnings = [];
  const latestByPatient = {};
  let setupIncomplete = false;
  const logContext = (label, details) => {
    if (typeof dashboardLogContext_ === 'function') {
      dashboardLogContext_(label, details);
    } else if (typeof dashboardWarn_ === 'function') {
      const payload = details ? ` ${details}` : '';
      dashboardWarn_(`[${label}]${payload}`);
    }
  };

  const wb = dashboardGetSpreadsheet_();
  if (!wb) {
    const warning = 'スプレッドシートを取得できませんでした';
    warnings.push(warning);
    setupIncomplete = true;
    dashboardWarn_('[loadNotes] spreadsheet unavailable');
    logContext('loadNotesUncached:done', `notes=0 warnings=${warnings.length} setupIncomplete=true`);
    return { notes: latestByPatient, warnings, setupIncomplete };
  }
  const sheetName = typeof DASHBOARD_SHEET_NOTES !== 'undefined' ? DASHBOARD_SHEET_NOTES : '申し送り';
  const sheet = wb && wb.getSheetByName ? wb.getSheetByName(sheetName) : null;
  if (!sheet) {
    const warning = `${sheetName}シートが見つかりません`;
    warnings.push(warning);
    setupIncomplete = true;
    dashboardWarn_(`[loadNotes] sheet not found: ${sheetName}`);
    logContext('loadNotesUncached:done', `notes=0 warnings=${warnings.length} setupIncomplete=true`);
    return { notes: latestByPatient, warnings, setupIncomplete };
  }

  const lastRow = sheet.getLastRow ? sheet.getLastRow() : 0;
  if (lastRow < 2) {
    logContext('loadNotesUncached:done', `notes=0 warnings=${warnings.length} setupIncomplete=${setupIncomplete} lastRow=${lastRow}`);
    return { notes: latestByPatient, warnings, setupIncomplete };
  }

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

  logContext('loadNotesUncached:done', `notes=${Object.keys(latestByPatient).length} warnings=${warnings.length} setupIncomplete=${setupIncomplete}`);
  return { notes: latestByPatient, warnings, setupIncomplete };
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
