/**
 * 患者情報シートを読み込み、患者IDを主キーとした情報と氏名マッピングを返す。
 */
function loadPatientInfo(options) {
  const opts = options || {};
  return loadPatientInfoUncached_(opts);
}

function loadPatientInfoUncached_(options) {
  const opts = options || {};
  const patients = {};
  const nameToId = {};
  const warnings = [];
  let setupIncomplete = false;
  const logContext = (label, details) => {
    if (typeof dashboardLogContext_ === 'function') {
      dashboardLogContext_(label, details);
    } else if (typeof dashboardWarn_ === 'function') {
      const payload = details ? ` ${details}` : '';
      dashboardWarn_(`[${label}]${payload}`);
    }
  };

  const wb = opts.dashboardSpreadsheet || null;
  if (!wb) {
    const warning = 'スプレッドシートを取得できませんでした';
    warnings.push(warning);
    setupIncomplete = true;
    dashboardWarn_('[loadPatientInfo] spreadsheet unavailable');
    logContext('loadPatientInfo:done', `patients=0 warnings=${warnings.length} setupIncomplete=true`);
    return { patients, nameToId, warnings, setupIncomplete };
  }
  const sheetName = typeof DASHBOARD_SHEET_PATIENTS !== 'undefined' ? DASHBOARD_SHEET_PATIENTS : '患者情報';
  logContext('loadPatientInfo:openSheet', `sheetName=${sheetName}`);
  const sheet = wb && wb.getSheetByName ? wb.getSheetByName(sheetName) : null;
  if (!sheet) {
    const warning = `${sheetName}シートが見つかりません`;
    warnings.push(warning);
    setupIncomplete = true;
    dashboardWarn_(`[loadPatientInfo] sheet not found: ${sheetName}`);
    logContext('loadPatientInfo:done', `patients=0 warnings=${warnings.length} setupIncomplete=true`);
    return { patients, nameToId, warnings, setupIncomplete };
  }

  const lastRow = sheet.getLastRow ? sheet.getLastRow() : 0;
  logContext('loadPatientInfo:rowCount', `sheetName=${sheetName} rowCount=${lastRow}`);
  if (lastRow < 2) {
    throw new Error(`Patient master empty. SheetName=${sheetName}, RowCount=${lastRow}, Header=[]`);
  }

  const lastCol = sheet.getLastColumn ? sheet.getLastColumn() : sheet.getMaxColumns ? sheet.getMaxColumns() : 0;
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  logContext('loadPatientInfo:header', `sheetName=${sheetName} header=${JSON.stringify(headers)}`);

  const colPid = dashboardResolveColumn_(headers, ['患者ID', 'patientId', 'ID', 'id', '施術録番号'], 1);
  logContext('loadPatientInfo:patientIdColumn', `index=${colPid}`);
  const colName = dashboardResolveColumn_(headers, ['氏名', '名前', '患者名'], 2);
  const colConsent = dashboardResolveColumn_(headers, ['同意期限', '同意書期限', '同意有効期限', '同意期限日'], 0);

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const displayValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const patientIdSamples = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i] || [];
    const rowDisplay = displayValues[i] || [];
    const rowNumber = i + 2;

    const patientId = dashboardNormalizePatientId_(rowDisplay[colPid - 1] || row[colPid - 1]);
    if (patientIdSamples.length < 5) {
      patientIdSamples.push(patientId);
    }
    const name = String(rowDisplay[colName - 1] || row[colName - 1] || '').trim();
    const consentExpiry = colConsent
      ? String(rowDisplay[colConsent - 1] || row[colConsent - 1] || '').trim()
      : '';

    if (!patientId) {
      warnings.push(`患者IDが空です (row:${rowNumber})`);
      dashboardWarn_(`[loadPatientInfo] missing patientId at row ${rowNumber}`);
      continue;
    }
    if (!name) {
      warnings.push(`氏名が未入力の患者があります (患者ID:${patientId})`);
      dashboardWarn_(`[loadPatientInfo] missing name for patientId ${patientId}`);
    }

    const normalizedNameKey = dashboardNormalizeNameKey_(name);
    if (normalizedNameKey && !nameToId[normalizedNameKey]) {
      nameToId[normalizedNameKey] = patientId;
    }

    const raw = {};
    headers.forEach((h, idx) => {
      const key = String(h || '').trim();
      if (!key) return;
      raw[key] = row[idx];
    });

    patients[patientId] = {
      patientId,
      name,
      consentExpiry,
      raw
    };
  }

  logContext('loadPatientInfo:patientIdSamples', JSON.stringify(patientIdSamples));
  const patientMapSize = Object.keys(patients).length;
  logContext('loadPatientInfo:patientMapSize', `size=${patientMapSize}`);
  if (patientMapSize === 0) {
    throw new Error(`Patient master empty. SheetName=${sheetName}, RowCount=${lastRow}, Header=${JSON.stringify(headers)}`);
  }

  logContext('loadPatientInfo:done', `patients=${patientMapSize} warnings=${warnings.length} setupIncomplete=${setupIncomplete}`);
  return { patients, nameToId, warnings, setupIncomplete };
}
