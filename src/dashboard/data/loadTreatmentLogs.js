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
  const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo({ dashboardSpreadsheet: opts.dashboardSpreadsheet }) : null);
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
  const logPerf = message => {
    if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
      Logger.log(message);
    } else if (typeof dashboardWarn_ === 'function') {
      dashboardWarn_(message);
    } else {
      logContext('perf', message);
    }
  };
  const perfBeforeStartedAt = Date.now();

  const wb = opts.dashboardSpreadsheet || null;
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

  const colTimestamp = dashboardResolveColumn_(headers, ['日時', '日付', 'timestamp', 'ts', 'タイムスタンプ'], 1);
  const colPatientId = dashboardResolveColumn_(headers, [
    '施術録番号', '施術録No', '施術録NO', '記録番号', 'カルテ番号', '患者ID', '患者番号', 'recNo', 'patientId', 'id'
  ], 2);
  const colPatientName = dashboardResolveColumn_(headers, ['氏名', '名前', '患者名'], colPatientId ? 0 : 2);
  const colCreatedBy = dashboardResolveColumn_(headers, ['作成者', '記録者', '入力者', '編集者', 'createdby', 'createdBy'], 0);
  const colStaffName = dashboardResolveColumn_(headers, ['施術者', '担当者', '担当', 'スタッフ名', 'staffName', 'staff'], 0);
  const colStaffEmail = dashboardResolveColumn_(headers, ['メール', '担当メール', 'email', 'mail', 'createdbyemail'], 0);
  const colStaffId = dashboardResolveColumn_(headers, ['担当者ID', 'スタッフID', 'staffId', 'staffid'], 0);
  const searchableColumns = [
    dashboardResolveColumn_(headers, ['施術内容', '内容', '記録', 'メモ', 'ノート', '備考', 'コメント', '対応内容', '申し送り', '自由記述', 'text', 'note', 'memo', 'comment', 'body', 'content'], 0),
    dashboardResolveColumn_(headers, ['SOAP', 'S', 'O', 'A', 'P'], 0)
  ].filter(function(col, index, arr) { return !!col && arr.indexOf(col) === index; });

  const tz = dashboardResolveTimeZone_();
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const monthStart = dashboardStartOfMonth_(tz, now);
  const previousMonthStart = new Date(monthStart.getTime());
  previousMonthStart.setMonth(previousMonthStart.getMonth() - 1);
  const monthEnd = new Date(monthStart.getTime());
  monthEnd.setMonth(monthEnd.getMonth() + 1);
  const prevMonthEnd = dashboardEndOfPreviousMonth_(monthStart);
  const allowedPatientIds = opts.allowedPatientIds instanceof Set
    ? new Set(Array.from(opts.allowedPatientIds).map(function(pid) { return dashboardNormalizePatientId_(pid); }).filter(function(pid) { return !!pid; }))
    : null;

  const headerSnapshot = {
    createdBy: colCreatedBy ? headers[colCreatedBy - 1] : '',
    staffName: colStaffName ? headers[colStaffName - 1] : '',
    staffEmail: colStaffEmail ? headers[colStaffEmail - 1] : '',
    staffId: colStaffId ? headers[colStaffId - 1] : ''
  };
  logContext('loadTreatmentLogs:staffColumns', JSON.stringify(headerSnapshot));

  if (allowedPatientIds && allowedPatientIds.size === 0) {
    logPerf('[perf] loadTreatmentLogsBefore=' + (Date.now() - perfBeforeStartedAt) + 'ms');
    logPerf('[perf] loadTreatmentLogsAfter=0ms');
    logPerf('[perf] rowsProcessed=0');
    logContext('loadTreatmentLogs:done', `logs=0 warnings=${warnings.length} setupIncomplete=${setupIncomplete}`);
    return { logs, warnings, lastStaffByPatient, setupIncomplete };
  }

  const dateValues = sheet.getRange(2, colTimestamp, lastRow - 1, 1).getValues();
  const dateDisplayValues = sheet.getRange(2, colTimestamp, lastRow - 1, 1).getDisplayValues();
  let startDataIndex = -1;
  let endDataIndex = -1;
  for (let i = 0; i < dateValues.length; i++) {
    const timestamp = dashboardParseTimestamp_(dateValues[i][0] || dateDisplayValues[i][0]);
    if (!timestamp) continue;
    if (timestamp < previousMonthStart || timestamp >= monthEnd) continue;
    if (startDataIndex < 0) startDataIndex = i;
    endDataIndex = i;
  }

  const perfBeforeDuration = Date.now() - perfBeforeStartedAt;
  logPerf('[perf] loadTreatmentLogsBefore=' + perfBeforeDuration + 'ms');

  if (startDataIndex < 0 || endDataIndex < startDataIndex) {
    logPerf('[perf] loadTreatmentLogsAfter=0ms');
    logPerf('[perf] rowsProcessed=0');
    logContext('loadTreatmentLogs:done', `logs=0 warnings=${warnings.length} setupIncomplete=${setupIncomplete}`);
    return { logs, warnings, lastStaffByPatient, setupIncomplete };
  }

  const dataStartRow = startDataIndex + 2;
  const dataRowCount = endDataIndex - startDataIndex + 1;
  const requiredColumns = [
    colTimestamp,
    colPatientId,
    colPatientName,
    colCreatedBy,
    colStaffName,
    colStaffEmail,
    colStaffId
  ].concat(searchableColumns)
    .filter(function(col, index, arr) { return !!col && arr.indexOf(col) === index; });

  const rowDataByColumn = {};
  for (let ci = 0; ci < requiredColumns.length; ci++) {
    const col = requiredColumns[ci];
    rowDataByColumn[col] = {
      values: sheet.getRange(dataStartRow, col, dataRowCount, 1).getValues(),
      displayValues: sheet.getRange(dataStartRow, col, dataRowCount, 1).getDisplayValues()
    };
  }

  let rowsProcessed = 0;
  const perfAfterStartedAt = Date.now();

  for (let i = 0; i < dataRowCount; i++) {
    rowsProcessed += 1;
    const rowNumber = dataStartRow + i;
    const readValue = function(col) {
      if (!col || !rowDataByColumn[col]) return '';
      return rowDataByColumn[col].values[i][0];
    };
    const readDisplay = function(col) {
      if (!col || !rowDataByColumn[col]) return '';
      return rowDataByColumn[col].displayValues[i][0];
    };

    const timestamp = dashboardParseTimestamp_(readValue(colTimestamp) || readDisplay(colTimestamp));
    if (!timestamp) {
      warnings.push(`施術日時を解釈できません (row:${rowNumber})`);
      continue;
    }
    if (timestamp < previousMonthStart || timestamp >= monthEnd) {
      continue;
    }

    const patientIdCell = readValue(colPatientId);
    const patientIdDisplay = readDisplay(colPatientId);
    const patientIdRaw = colPatientId
      ? (dashboardNormalizePatientId_(patientIdCell) || dashboardNormalizePatientId_(patientIdDisplay))
      : '';
    const patientName = String((colPatientName && (readDisplay(colPatientName) || readValue(colPatientName))) || '').trim();
    const mappedPatientId = patientIdRaw || dashboardResolvePatientIdFromName_(patientName, nameToId);

    const normalizedPatientId = mappedPatientId ? dashboardNormalizePatientId_(mappedPatientId) : '';
    if (!normalizedPatientId) {
      logContext('loadTreatmentLogs:skipMissingPatientId', `row=${rowNumber}`);
      continue;
    }
    if (allowedPatientIds && !allowedPatientIds.has(normalizedPatientId)) {
      continue;
    }
    if (!Object.prototype.hasOwnProperty.call(patients, normalizedPatientId)) {
      const samplePatientIds = Object.keys(patients).slice(0, 10);
      const rawType = patientIdCell === null ? 'null' : typeof patientIdCell;
      const rawValue = String(patientIdCell === undefined ? '' : patientIdCell);
      logContext('loadTreatmentLogs:unknownPatientDebug', JSON.stringify({
        row: rowNumber,
        patientMapSampleKeys: samplePatientIds,
        rawPatientId: {
          type: rawType,
          value: rawValue
        },
        normalizedPatientId
      }));
      logContext('loadTreatmentLogs:skipUnknownPatientId', `row=${rowNumber} patientId=${normalizedPatientId}`);
      continue;
    }

    const createdByRaw = colCreatedBy ? String(readDisplay(colCreatedBy) || readValue(colCreatedBy) || '').trim() : '';
    const staffNameRaw = colStaffName ? String(readDisplay(colStaffName) || readValue(colStaffName) || '').trim() : '';
    const staffEmailRaw = colStaffEmail ? String(readDisplay(colStaffEmail) || readValue(colStaffEmail) || '').trim() : '';
    const staffIdRaw = colStaffId ? String(readDisplay(colStaffId) || readValue(colStaffId) || '').trim() : '';
    const dateKey = dashboardFormatDate_(timestamp, tz, 'yyyy-MM-dd');
    const searchText = searchableColumns
      .map(function(col) {
        return String(readDisplay(col) || readValue(col) || '').trim();
      })
      .filter(function(text) { return !!text; })
      .join('\n');

    if (i < 20) {
      logContext('loadTreatmentLogs:staffValueSample', JSON.stringify({
        row: rowNumber,
        createdByRaw,
        staffNameRaw,
        staffEmailRaw,
        staffIdRaw
      }));
    }

    const entry = {
      row: rowNumber,
      patientId: normalizedPatientId,
      patientName,
      createdByEmail: staffEmailRaw || createdByRaw,
      staffName: staffNameRaw,
      staffId: staffIdRaw,
      staffKeys: {
        email: dashboardNormalizeStaffKey_(staffEmailRaw),
        name: dashboardNormalizeStaffKey_(staffNameRaw || createdByRaw),
        staffId: dashboardNormalizeStaffKey_(staffIdRaw)
      },
      timestamp,
      dateKey,
      searchText
    };

    logs.push(entry);

    if (timestamp <= prevMonthEnd) {
      const prev = lastStaffByPatient[normalizedPatientId];
      if (!prev || (prev.timestamp && prev.timestamp < timestamp)) {
        lastStaffByPatient[normalizedPatientId] = { email: entry.createdByEmail, timestamp };
      }
    }
  }

  Object.keys(lastStaffByPatient).forEach(pid => {
    const entry = lastStaffByPatient[pid];
    lastStaffByPatient[pid] = entry ? entry.email || '' : '';
  });

  const perfAfterDuration = Date.now() - perfAfterStartedAt;
  logPerf('[perf] loadTreatmentLogsAfter=' + perfAfterDuration + 'ms');
  logPerf('[perf] rowsProcessed=' + rowsProcessed);

  logContext('loadTreatmentLogs:done', `logs=${logs.length} warnings=${warnings.length} setupIncomplete=${setupIncomplete}`);
  return { logs, warnings, lastStaffByPatient, setupIncomplete };
}
