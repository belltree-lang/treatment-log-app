/**
 * 施術録を読み込み、患者名を患者IDへマッピングした結果を返す。
 */
function loadTreatmentLogs(options) {
  const opts = options || {};
  const fetchFn = () => loadTreatmentLogsUncached_(opts);
  if (typeof dashboardCacheFetch_ === 'function') {
    return dashboardCacheFetch_(dashboardCacheKey_('treatmentLogs:v1'), fetchFn, DASHBOARD_CACHE_TTL_SECONDS, opts);
  }
  return fetchFn();
}

function loadTreatmentLogsUncached_(options) {
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
