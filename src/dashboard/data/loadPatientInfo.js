/**
 * 患者情報シートを読み込み、患者IDを主キーとした情報と氏名マッピングを返す。
 */
function loadPatientInfo(options) {
  const opts = options || {};
  const fetchFn = () => loadPatientInfoUncached_(opts);
  if (opts && opts.cache === false) return fetchFn();
  return dashboardCacheFetch_(dashboardCacheKey_('patientInfo:v1'), fetchFn, DASHBOARD_CACHE_TTL_SECONDS);
}

function loadPatientInfoUncached_(_options) {
  const patients = {};
  const nameToId = {};
  const warnings = [];

  const wb = dashboardGetSpreadsheet_();
  const sheet = wb && wb.getSheetByName ? wb.getSheetByName('患者情報') : null;
  if (!sheet) {
    warnings.push('患者情報シートが見つかりません');
    dashboardWarn_('[loadPatientInfo] sheet not found');
    return { patients, nameToId, warnings };
  }

  const lastRow = sheet.getLastRow ? sheet.getLastRow() : 0;
  if (lastRow < 2) return { patients, nameToId, warnings };

  const lastCol = sheet.getLastColumn ? sheet.getLastColumn() : sheet.getMaxColumns ? sheet.getMaxColumns() : 0;
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const displayValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  const colPid = dashboardResolveColumn_(headers, ['患者ID', 'patientId', 'ID', 'id', '施術録番号'], 1);
  const colName = dashboardResolveColumn_(headers, ['氏名', '名前', '患者名'], 2);
  const colConsent = dashboardResolveColumn_(headers, ['同意期限', '同意書期限', '同意有効期限', '同意期限日'], 0);

  for (let i = 0; i < values.length; i++) {
    const row = values[i] || [];
    const rowDisplay = displayValues[i] || [];
    const rowNumber = i + 2;

    const patientId = dashboardNormalizePatientId_(rowDisplay[colPid - 1] || row[colPid - 1]);
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

  return { patients, nameToId, warnings };
}
