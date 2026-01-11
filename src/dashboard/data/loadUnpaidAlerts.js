/**
 * 未回収履歴から連続未回収アラートを組み立てる。
 * @param {Object} [options]
 * @param {Object} [options.patientInfo]
 * @param {number} [options.consecutiveMonths]
 * @param {Date} [options.now]
 * @return {{alerts: Object[], warnings: string[], setupIncomplete: boolean}}
 */
function loadUnpaidAlerts(options) {
  const opts = options || {};
  const normalizedThreshold = normalizeUnpaidThreshold_(opts.consecutiveMonths);
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const fetchOptions = Object.assign({}, opts, { consecutiveMonths: normalizedThreshold, now });
  return loadUnpaidAlertsUncached_(fetchOptions);
}

function loadUnpaidAlertsUncached_(options) {
  const opts = options || {};
  const threshold = normalizeUnpaidThreshold_(opts.consecutiveMonths);
  const tz = dashboardResolveTimeZone_();

  const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo(opts) : null);
  const patients = patientInfo && patientInfo.patients ? patientInfo.patients : {};
  const nameToId = patientInfo && patientInfo.nameToId ? patientInfo.nameToId : {};

  const history = readUnpaidHistory_(Object.assign({}, opts, { tz, nameToId }));

  const warnings = [];
  if (patientInfo && Array.isArray(patientInfo.warnings)) warnings.push.apply(warnings, patientInfo.warnings);
  if (history && Array.isArray(history.warnings)) warnings.push.apply(warnings, history.warnings);

  const alerts = buildUnpaidAlerts_(history.entries, patients, threshold, tz);

  return {
    alerts,
    warnings,
    setupIncomplete: !!(patientInfo && patientInfo.setupIncomplete) || !!(history && history.setupIncomplete)
  };
}

function readUnpaidHistory_(options) {
  const opts = options || {};
  const entries = [];
  const warnings = [];
  let setupIncomplete = false;

  const wb = dashboardGetSpreadsheet_();
  if (!wb) {
    warnings.push('スプレッドシートを取得できませんでした');
    setupIncomplete = true;
    return { entries, warnings, setupIncomplete };
  }

  const sheetName = typeof DASHBOARD_SHEET_UNPAID_HISTORY !== 'undefined' ? DASHBOARD_SHEET_UNPAID_HISTORY : '未回収履歴';
  const sheet = wb && wb.getSheetByName ? wb.getSheetByName(sheetName) : null;
  if (!sheet) {
    warnings.push(`${sheetName}シートが見つかりません`);
    setupIncomplete = true;
    return { entries, warnings, setupIncomplete };
  }

  const lastRow = sheet.getLastRow ? sheet.getLastRow() : 0;
  if (lastRow < 2) return { entries, warnings, setupIncomplete };

  const lastCol = sheet.getLastColumn ? sheet.getLastColumn() : sheet.getMaxColumns ? sheet.getMaxColumns() : 0;
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const displayValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  const colPatientId = dashboardResolveColumn_(headers, ['患者ID', 'patientId', 'ID', 'id'], 1);
  const colPatientName = dashboardResolveColumn_(headers, ['氏名', '名前', '患者名'], 0);
  const colMonth = dashboardResolveColumn_(headers, ['対象月', '月', '年月', '請求月'], 2);
  const colAmount = dashboardResolveColumn_(headers, ['金額', '未回収金額', '請求額', '額'], 3);
  const colReason = dashboardResolveColumn_(headers, ['理由', '未回収理由', '未回収備考'], 4);
  const colMemo = dashboardResolveColumn_(headers, ['備考', 'メモ'], 5);
  const colRecordedAt = dashboardResolveColumn_(headers, ['記録日時', 'timestamp', '記録日', '入力日時'], 6);

  for (let i = 0; i < values.length; i++) {
    const row = values[i] || [];
    const rowDisplay = displayValues[i] || [];
    const rowNumber = i + 2;

    let patientId = dashboardNormalizePatientId_(rowDisplay[colPatientId - 1] || row[colPatientId - 1]);
    if (!patientId && colPatientName) {
      const name = rowDisplay[colPatientName - 1] || row[colPatientName - 1];
      patientId = dashboardResolvePatientIdFromName_(name, opts.nameToId);
    }
    if (!patientId) {
      warnings.push(`未回収履歴の患者IDが空です (row:${rowNumber})`);
      continue;
    }

    const monthKey = normalizeUnpaidMonthKey_(rowDisplay[colMonth - 1] || row[colMonth - 1], opts.tz);
    if (!monthKey) {
      warnings.push(`未回収履歴の対象月を解釈できません (row:${rowNumber})`);
      continue;
    }

    const amount = coerceUnpaidAmount_(row[colAmount - 1] != null ? row[colAmount - 1] : rowDisplay[colAmount - 1]);
    const reason = String(rowDisplay[colReason - 1] || row[colReason - 1] || '').trim();
    const memo = String(rowDisplay[colMemo - 1] || row[colMemo - 1] || '').trim();
    const recordedAt = dashboardParseTimestamp_(row[colRecordedAt - 1] || rowDisplay[colRecordedAt - 1]);

    entries.push({ patientId, monthKey, amount, reason, memo, recordedAt });
  }

  return { entries, warnings, setupIncomplete };
}

function buildUnpaidAlerts_(entries, patients, threshold, tz) {
  const grouped = {};
  (entries || []).forEach(entry => {
    const pid = dashboardNormalizePatientId_(entry && entry.patientId);
    if (!pid || !entry || !entry.monthKey) return;
    if (!grouped[pid]) {
      grouped[pid] = { totals: {}, records: {} };
    }
    grouped[pid].totals[entry.monthKey] = (grouped[pid].totals[entry.monthKey] || 0) + (entry.amount || 0);
    if (!grouped[pid].records[entry.monthKey]) grouped[pid].records[entry.monthKey] = [];
    grouped[pid].records[entry.monthKey].push({
      amount: entry.amount || 0,
      reason: entry.reason || '',
      memo: entry.memo || '',
      recordedAt: entry.recordedAt instanceof Date && !Number.isNaN(entry.recordedAt.getTime())
        ? dashboardFormatDate_(entry.recordedAt, tz, 'yyyy-MM-dd HH:mm')
        : ''
    });
  });

  const alerts = [];
  Object.keys(grouped).forEach(pid => {
    const bucket = grouped[pid];
    const runKeys = summarizeLatestRun_(bucket.totals);
    if (runKeys.length < threshold) return;

    const months = runKeys.map(key => ({
      key,
      amount: bucket.totals[key] || 0,
      records: (bucket.records[key] || []).map(record => Object.assign({}, record)),
      followUp: { phone: false, visit: false }
    }));

    const totalAmount = months.reduce((sum, month) => sum + (Number(month.amount) || 0), 0);
    const patient = patients && patients[pid] ? patients[pid] : {};

    alerts.push({
      patientId: pid,
      patientName: patient.name || patient.patientName || '',
      consecutiveMonths: runKeys.length,
      totalAmount,
      months,
      followUp: { phone: false, visit: false }
    });
  });

  alerts.sort((a, b) => {
    if (b.consecutiveMonths !== a.consecutiveMonths) return b.consecutiveMonths - a.consecutiveMonths;
    if (b.totalAmount !== a.totalAmount) return b.totalAmount - a.totalAmount;
    return (a.patientName || '').localeCompare(b.patientName || '', 'ja');
  });

  return alerts;
}

function summarizeLatestRun_(monthTotals) {
  const parsed = Object.keys(monthTotals || {})
    .map(key => ({ key, ordinal: unpaidMonthOrdinal_(key) }))
    .filter(entry => entry.ordinal !== null)
    .sort((a, b) => b.ordinal - a.ordinal);

  if (!parsed.length) return [];

  const run = [parsed[0].key];
  for (let i = 1; i < parsed.length; i++) {
    if (parsed[i - 1].ordinal - parsed[i].ordinal === 1) {
      run.push(parsed[i].key);
    } else {
      break;
    }
  }
  return run;
}

function normalizeUnpaidThreshold_(value) {
  const raw = Number(value);
  if (Number.isFinite(raw) && raw > 0) return Math.floor(raw);
  if (typeof DASHBOARD_UNPAID_ALERT_MONTHS === 'number' && Number.isFinite(DASHBOARD_UNPAID_ALERT_MONTHS)) {
    return Math.max(1, Math.floor(DASHBOARD_UNPAID_ALERT_MONTHS));
  }
  return 3;
}

function normalizeUnpaidMonthKey_(value, tz) {
  const parsed = dashboardCoerceDate_(value);
  if (parsed instanceof Date && !Number.isNaN(parsed.getTime())) {
    return dashboardFormatDate_(new Date(parsed.getFullYear(), parsed.getMonth(), 1), tz, 'yyyy-MM');
  }
  const str = String(value == null ? '' : value).trim();
  if (!str) return '';
  const digits = str.replace(/[^0-9]/g, '');
  if (digits.length >= 6) {
    const year = digits.slice(0, 4);
    const month = digits.slice(4, 6);
    if (Number(month) >= 1 && Number(month) <= 12) {
      return `${year}-${month}`;
    }
  }
  return '';
}

function unpaidMonthOrdinal_(monthKey) {
  const match = String(monthKey || '').match(/^([0-9]{4})-([0-9]{2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (!Number.isFinite(year) || !Number.isFinite(month)) return null;
  return year * 12 + (month - 1);
}

function coerceUnpaidAmount_(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const digits = String(value == null ? '' : value).replace(/[^0-9.-]/g, '');
  const parsed = Number(digits);
  return Number.isFinite(parsed) ? parsed : 0;
}
