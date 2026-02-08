/**
 * 施術録を読み込み、患者名を患者IDへマッピングした結果を返す。
 */
function loadTreatmentLogs(options) {
  const opts = options || {};
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const fetchOptions = Object.assign({}, opts, { now });
  return loadTreatmentLogsUncached_(fetchOptions);
}

function loadTreatmentLogsUncached_(options) {
  const opts = options || {};
  const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo() : null);
  const nameToId = opts.nameToId || (patientInfo && patientInfo.nameToId) || {};
  const warnings = patientInfo && Array.isArray(patientInfo.warnings) ? [].concat(patientInfo.warnings) : [];
  let setupIncomplete = !!(patientInfo && patientInfo.setupIncomplete);
  const patients = patientInfo && patientInfo.patients ? patientInfo.patients : {};
  const logs = [];
  const lastStaffByPatient = {};
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
    warnings.push('スプレッドシートを取得できませんでした');
    setupIncomplete = true;
    dashboardWarn_('[loadTreatmentLogs] spreadsheet unavailable');
    logContext('loadTreatmentLogs:done', `logs=0 warnings=${warnings.length} setupIncomplete=true`);
    return { logs, warnings, lastStaffByPatient, setupIncomplete };
  }
  const sheetName = typeof DASHBOARD_SHEET_TREATMENTS !== 'undefined' ? DASHBOARD_SHEET_TREATMENTS : '施術録';
  const sheet = wb && wb.getSheetByName ? wb.getSheetByName(sheetName) : null;
  if (!sheet) {
    const warning = `${sheetName}シートが見つかりません`;
    warnings.push(warning);
    setupIncomplete = true;
    dashboardWarn_(`[loadTreatmentLogs] sheet not found: ${sheetName}`);
    logContext('loadTreatmentLogs:done', `logs=0 warnings=${warnings.length} setupIncomplete=true`);
    return { logs, warnings, lastStaffByPatient, setupIncomplete };
  }

  const lastRow = sheet.getLastRow ? sheet.getLastRow() : 0;
  if (lastRow < 2) {
    logContext('loadTreatmentLogs:done', `logs=0 warnings=${warnings.length} setupIncomplete=${setupIncomplete} lastRow=${lastRow}`);
    return { logs, warnings, lastStaffByPatient, setupIncomplete };
  }

  const lastCol = sheet.getLastColumn ? sheet.getLastColumn() : sheet.getMaxColumns ? sheet.getMaxColumns() : 0;
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const displayValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  const colTimestamp = dashboardResolveColumn_(headers, ['日時', '日付', 'timestamp', 'ts', 'タイムスタンプ'], 1);
  const colPatientId = dashboardResolveColumn_(headers, ['患者ID', 'patientId', 'id'], 0);
  const colPatientName = dashboardResolveColumn_(headers, ['氏名', '名前', '患者名'], colPatientId ? 0 : 2);
  const colEmail = dashboardResolveColumn_(headers, ['作成者', '担当者', 'email', 'createdbyemail'], 0);

  const tz = dashboardResolveTimeZone_();
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const monthStart = dashboardStartOfMonth_(tz, now);
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

    const normalizedPatientId = mappedPatientId ? dashboardNormalizePatientId_(mappedPatientId) : '';
    if (!normalizedPatientId) {
      logContext('loadTreatmentLogs:skipMissingPatientId', `row=${rowNumber}`);
      continue;
    }
    if (!Object.prototype.hasOwnProperty.call(patients, normalizedPatientId)) {
      logContext('loadTreatmentLogs:skipUnknownPatientId', `row=${rowNumber} patientId=${normalizedPatientId}`);
      continue;
    }

    const entry = {
      row: rowNumber,
      patientId: normalizedPatientId,
      patientName,
      createdByEmail,
      timestamp,
      dateKey
    };

    logs.push(entry);

    if (timestamp <= prevMonthEnd) {
      const prev = lastStaffByPatient[normalizedPatientId];
      if (!prev || (prev.timestamp && prev.timestamp < timestamp)) {
        lastStaffByPatient[normalizedPatientId] = { email: createdByEmail, timestamp };
      }
    }
  }

  Object.keys(lastStaffByPatient).forEach(pid => {
    const entry = lastStaffByPatient[pid];
    lastStaffByPatient[pid] = entry ? entry.email || '' : '';
  });

  logContext('loadTreatmentLogs:done', `logs=${logs.length} warnings=${warnings.length} setupIncomplete=${setupIncomplete}`);
  return { logs, warnings, lastStaffByPatient, setupIncomplete };
}
