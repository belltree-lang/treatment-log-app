function doGet(e) {
  e = e || {};

  const dashboardResponse = handleDashboardDoGet_(e);
  if (dashboardResponse) {
    return dashboardResponse;
  }

  const view = e.parameter ? (e.parameter.view || 'billing') : 'billing';
  let templateFile = '';

  switch (view) {
    case 'billing':      templateFile = 'billing'; break;
    case 'dashboard':    templateFile = 'dashboard'; break;
    default:             templateFile = 'billing'; break;
  }

  const template = HtmlService.createTemplateFromFile(templateFile);
  template.baseUrl = ScriptApp.getService().getUrl() || '';
  template.patientId = e.parameter && e.parameter.id ? e.parameter.id : '';
  template.payrollPdfData = {};

  if (e.parameter && e.parameter.lead) template.lead = e.parameter.lead;

  return template
    .evaluate()
    .setTitle(templateFile === 'dashboard' ? '施術者ダッシュボード' : '請求処理アプリ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function handleDashboardDoGet_(e) {
  if (!shouldHandleDashboardRequest_(e)) return null;

  if (shouldHandleDashboardApi_(e)) {
    const data = typeof getDashboardData === 'function' ? getDashboardData() : {};
    return createJsonResponse_(data);
  }

  const template = HtmlService.createTemplateFromFile('dashboard');
  template.baseUrl = ScriptApp.getService().getUrl() || '';

  return template
    .evaluate()
    .setTitle('施術者ダッシュボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function shouldHandleDashboardRequest_(e) {
  const path = (e && e.pathInfo ? String(e.pathInfo) : '').replace(/^\/+|\/+$/g, '').toLowerCase();
  if (path === 'dashboard' || path === 'getdashboarddata') return true;
  const view = e && e.parameter ? String(e.parameter.view || '').toLowerCase() : '';
  if (view === 'dashboard') return true;
  const action = e && e.parameter ? String(e.parameter.action || e.parameter.api || '').toLowerCase() : '';
  return action === 'getdashboarddata';
}

function shouldHandleDashboardApi_(e) {
  const path = (e && e.pathInfo ? String(e.pathInfo) : '').replace(/^\/+|\/+$/g, '').toLowerCase();
  if (path === 'getdashboarddata') return true;
  const action = e && e.parameter ? (e.parameter.action || e.parameter.api) : '';
  return String(action || '').toLowerCase() === 'getdashboarddata';
}

function createJsonResponse_(payload) {
  if (typeof ContentService === 'undefined' || !ContentService || typeof ContentService.createTextOutput !== 'function') {
    return JSON.stringify(payload || {});
  }
  const output = ContentService.createTextOutput(JSON.stringify(payload || {}));
  if (output && typeof output.setMimeType === 'function' && ContentService && ContentService.MimeType) {
    output.setMimeType(ContentService.MimeType.JSON);
  }
  return output;
}

/**
 * Return the active spreadsheet (utility retained for compatibility).
 * @return {SpreadsheetApp.Spreadsheet}
 */
function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Resolve source data for billing generation (patients, visits, bank statuses).
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @return {Object} normalized source data including month metadata.
 */
function getBillingSource(billingMonth) {
  return getBillingSourceData(billingMonth);
}

/**
 * Expose billing row total calculation for the web UI.
 *
 * This delegates to the shared billingLogic implementation so that
 * browser previews stay in sync with the final server-side totals
 * used for PDFs and exports.
 * @param {Object} row - Merged billing row (base + edits) from the UI.
 * @return {Object} normalized totals for preview rendering.
 */
function calculateBillingRowTotalsServer(row) {
  const source = row && typeof row === 'object' ? row : {};
  const amountCalc = calculateBillingAmounts_({
    visitCount: source.visitCount,
    insuranceType: source.insuranceType,
    burdenRate: source.burdenRate,
    manualUnitPrice: source.manualUnitPrice != null ? source.manualUnitPrice : source.unitPrice,
    manualTransportAmount: Object.prototype.hasOwnProperty.call(source, 'manualTransportAmount')
      ? source.manualTransportAmount
      : source.transportAmount,
    unitPrice: source.unitPrice,
    medicalAssistance: source.medicalAssistance,
    carryOverAmount: source.carryOverAmount,
    selfPayItems: source.selfPayItems,
    manualSelfPayAmount: source.manualSelfPayAmount
  });

  return {
    visitCount: amountCalc.visits,
    treatmentUnitPrice: amountCalc.unitPrice,
    treatmentAmount: amountCalc.treatmentAmount,
    transportAmount: amountCalc.transportAmount,
    carryOverAmount: amountCalc.carryOverAmount,
    billingAmount: amountCalc.billingAmount,
    manualSelfPayAmount: amountCalc.manualSelfPayAmount,
    grandTotal: amountCalc.grandTotal
  };
}

const BILLING_CACHE_PREFIX = 'billing_prepared_';
const BILLING_CACHE_TTL_SECONDS = 3600; // 1 hour
const PREPARED_BILLING_SCHEMA_VERSION = 2;

function normalizeBillingMonthKeySafe_(value) {
  const candidates = [];
  if (value && typeof value === 'object') {
    if (value.key) candidates.push(value.key);
    if (value.billingMonth) candidates.push(value.billingMonth);
    if (value.ym) candidates.push(value.ym);
    if (value.month && value.month.key) candidates.push(value.month.key);
    if (value.month && value.month.ym) candidates.push(value.month.ym);
    if (value.month) candidates.push(value.month);
  } else {
    candidates.push(value);
  }

  for (let i = 0; i < candidates.length; i++) {
    const candidate = candidates[i];
    if (!candidate) continue;
    try {
      return normalizeBillingMonthInput(candidate).key;
    } catch (err) {
      const fallback = String(candidate || '').trim();
      if (fallback) return fallback;
    }
  }
  return '';
}

function buildBillingCacheKey_(billingMonthKey) {
  const monthKey = String(billingMonthKey || '').trim();
  if (!monthKey) return '';
  return BILLING_CACHE_PREFIX + monthKey;
}

function getBillingCache_() {
  try {
    return CacheService.getScriptCache();
  } catch (err) {
    console.warn('[billing] CacheService unavailable', err);
    return null;
  }
}

function buildStaffByPatient_() {
  const logs = loadTreatmentLogs_();
  const staffHistoryByPatient = {};
  const debug = { totalLogs: logs.length, missingPatientId: 0, missingStaff: 0 };

  logs.forEach(log => {
    const pid = log && log.patientId ? billingNormalizePatientId_(log.patientId) : '';
    const ts = log && log.timestamp;
    if (!pid) {
      debug.missingPatientId += 1;
      return;
    }
    if (!(ts instanceof Date) || isNaN(ts.getTime())) {
      return;
    }
    const staffKey = log && (log.createdByKey || billingNormalizeStaffKey_(log.createdByEmail)) || '';
    if (!staffKey) {
      debug.missingStaff += 1;
      return;
    }

    if (!staffHistoryByPatient[pid]) {
      staffHistoryByPatient[pid] = {};
    }
    const existing = staffHistoryByPatient[pid][staffKey];
    if (!existing || !existing.timestamp || ts > existing.timestamp) {
      staffHistoryByPatient[pid][staffKey] = { key: staffKey, email: log.createdByEmail, timestamp: ts };
    }
  });

  const staffByPatient = Object.keys(staffHistoryByPatient).reduce((map, pid) => {
    const staffEntries = staffHistoryByPatient[pid];
    const sorted = Object.keys(staffEntries)
      .map(key => staffEntries[key])
      .sort((a, b) => {
        const aTime = a.timestamp instanceof Date ? a.timestamp.getTime() : 0;
        const bTime = b.timestamp instanceof Date ? b.timestamp.getTime() : 0;
        return bTime - aTime;
      })
      .map(entry => entry.email || entry.key)
      .filter(email => !!email);
    map[pid] = sorted;
    return map;
  }, {});

  const staffDirectory = loadBillingStaffDirectory_();
  const staffDisplayByPatient = buildStaffDisplayByPatient_(staffByPatient, staffDirectory);
  billingLogger_.log('[billing] buildStaffByPatient_ summary=' + JSON.stringify({
    totalLogs: debug.totalLogs,
    missingPatientId: debug.missingPatientId,
    missingStaff: debug.missingStaff,
    staffByPatientSize: Object.keys(staffByPatient || {}).length,
    staffDirectorySize: Object.keys(staffDirectory || {}).length
  }));
  return { staffByPatient, staffDirectory, staffDisplayByPatient, staffHistoryByPatient };
}

function buildStaffDisplayByPatient_(staffByPatient, staffDirectory) {
  const result = {};
  const directory = staffDirectory || {};
  const displayLog = [];
  Object.keys(staffByPatient || {}).forEach(pid => {
    const emails = Array.isArray(staffByPatient[pid]) ? staffByPatient[pid] : [staffByPatient[pid]];
    const seen = new Set();
    const names = [];
    emails.forEach(email => {
      const key = billingNormalizeStaffKey_(email);
      if (!key || seen.has(key)) return;
      seen.add(key);
      const resolved = directory[key] || '';
      names.push(resolved || email || '');

      if (displayLog.length < 200) {
        displayLog.push({
          patientId: pid,
          email,
          normalizedKey: key,
          matched: !!directory[key],
          resolvedName: resolved || ''
        });
      }
    });
    result[pid] = names.filter(Boolean);
  });
  if (displayLog.length) {
    billingLogger_.log('[billing] buildStaffDisplayByPatient_: resolved staff detail=' + JSON.stringify(displayLog));
  }
  return result;
}

function clearBillingCache_(key) {
  if (!key) return;
  const cache = getBillingCache_();
  if (!cache || typeof cache.remove !== 'function') return;
  try {
    cache.remove(key);
  } catch (err) {
    console.warn('[billing] Failed to clear prepared cache', err);
  }
}

function ensurePreparedBillingSheet_() {
  const SHEET_NAME = 'PreparedBilling';
  const HEADER = ['billingMonth', 'preparedAt', 'preparedBy', 'payloadVersion', 'payloadJson', 'note'];
  const workbook = ss();
  let sheet = workbook.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }
  return sheet;
}

function savePreparedBillingToSheet_(billingMonth, preparedPayload) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth || (preparedPayload && preparedPayload.billingMonth));
  const normalized = normalizePreparedBilling_(preparedPayload);
  if (!monthKey || !normalized) {
    billingLogger_.log('[billing] savePreparedBillingToSheet_ skipped due to invalid payload');
    return null;
  }

  const payload = Object.assign({}, normalized, { billingMonth: monthKey });
  const sheet = ensurePreparedBillingSheet_();
  const preparedAtValue = (() => {
    const parsed = payload.preparedAt ? new Date(payload.preparedAt) : null;
    return parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : new Date();
  })();
  const preparedBy = (() => {
    try {
      const active = Session.getActiveUser && Session.getActiveUser();
      return active && typeof active.getEmail === 'function' ? active.getEmail() : '';
    } catch (err) {
      return '';
    }
  })();
  const payloadVersion = payload.schemaVersion || PREPARED_BILLING_SCHEMA_VERSION;
  const payloadJson = JSON.stringify(payload);
  const rowValues = [monthKey, preparedAtValue, preparedBy, payloadVersion, payloadJson, ''];

  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(row => String(row[0] || '').trim())
    : [];
  const existingIndex = existing.findIndex(value => value === monthKey);

  if (existingIndex >= 0) {
    const targetRow = existingIndex + 2;
    sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    return { billingMonth: monthKey, row: targetRow, updated: true };
  }

  sheet.insertRows(2, 1);
  sheet.getRange(2, 1, 1, rowValues.length).setValues([rowValues]);
  return { billingMonth: monthKey, row: 2, updated: false };
}

function loadPreparedBillingFromSheet_(billingMonth) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const workbook = ss();
  const sheet = workbook.getSheetByName('PreparedBilling');
  if (!monthKey || !sheet) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const target = values.find(row => String(row[0] || '').trim() === monthKey);
  if (!target) return null;

  const preparedAtCell = target[1];
  const preparedByCell = target[2];
  const payloadVersion = target[3];
  const payloadJson = target[4];

  if (!payloadJson) return null;
  try {
    const parsed = JSON.parse(payloadJson);
    const normalized = normalizePreparedBilling_(Object.assign({}, parsed, { billingMonth: monthKey }));
    if (!normalized) return null;
    const preparedAt = normalized.preparedAt || (preparedAtCell instanceof Date ? preparedAtCell.toISOString() : null);
    return Object.assign({}, normalized, {
      preparedAt,
      preparedBy: normalized.preparedBy || preparedByCell || '',
      payloadVersion: payloadVersion || normalized.schemaVersion || null
    });
  } catch (err) {
    billingLogger_.log('[billing] loadPreparedBillingFromSheet_ failed to parse payloadJson for ' + monthKey + ': ' + err);
    return null;
  }
}

function validatePreparedBillingPayload_(payload, expectedMonthKey) {
  if (!payload || typeof payload !== 'object') return { ok: false, reason: 'payload missing' };
  const billingMonth = payload.billingMonth || payload.month || '';
  if (!billingMonth) return { ok: false, reason: 'billingMonth missing' };
  if (expectedMonthKey && String(billingMonth) !== String(expectedMonthKey)) {
    return { ok: false, reason: 'billingMonth mismatch' };
  }

  const schemaVersion = payload.schemaVersion;
  if (schemaVersion == null) return { ok: false, reason: 'schemaVersion missing' };
  if (Number(schemaVersion) !== PREPARED_BILLING_SCHEMA_VERSION) {
    return { ok: false, reason: 'schemaVersion mismatch' };
  }

  const requiredArrays = [
    { key: 'billingJson', reason: 'billingJson missing' },
    { key: 'carryOverLedger', reason: 'carryOverLedger missing' },
    { key: 'unpaidHistory', reason: 'unpaidHistory missing' }
  ];
  const requiredMaps = [
    { key: 'carryOverLedgerMeta', reason: 'carryOverLedgerMeta missing' },
    { key: 'carryOverLedgerByPatient', reason: 'carryOverLedgerByPatient missing' },
    { key: 'visitsByPatient', reason: 'visitsByPatient missing' },
    { key: 'totalsByPatient', reason: 'totalsByPatient missing' },
    { key: 'patients', reason: 'patients missing' },
    { key: 'bankInfoByName', reason: 'bankInfoByName missing' },
    { key: 'staffByPatient', reason: 'staffByPatient missing' },
    { key: 'staffDirectory', reason: 'staffDirectory missing' },
    { key: 'staffDisplayByPatient', reason: 'staffDisplayByPatient missing' },
    { key: 'billingOverrideFlags', reason: 'billingOverrideFlags missing' },
    { key: 'carryOverByPatient', reason: 'carryOverByPatient missing' },
    { key: 'bankAccountInfoByPatient', reason: 'bankAccountInfoByPatient missing' }
  ];

  for (const field of requiredArrays) {
    if (!Array.isArray(payload[field.key])) return { ok: false, reason: field.reason };
  }

  for (const field of requiredMaps) {
    const value = payload[field.key];
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      return { ok: false, reason: field.reason };
    }
  }

  return { ok: true, billingMonth };
}

function loadPreparedBilling_(billingMonthKey, options) {
  const opts = options || {};
  const expectedMonthKey = normalizeBillingMonthKeySafe_(billingMonthKey);
  const key = buildBillingCacheKey_(expectedMonthKey);
  const withValidation = opts.withValidation === true;
  const allowInvalid = opts.allowInvalid === true;
  const wrapResult = (payload, validation) => withValidation ? { prepared: payload, validation: validation || null } : payload;
  if (!key) return wrapResult(null, { ok: false, reason: 'cache key missing' });
  const cache = getBillingCache_();
  if (!cache) return wrapResult(null, { ok: false, reason: 'cache unavailable' });
  const cached = cache.get(key);
  if (!cached) {
    try {
      Logger.log('[billing] loadPreparedBilling_: cache miss for ' + key);
    } catch (err) {
      // ignore logging errors in non-GAS environments
    }
    return wrapResult(null, { ok: false, reason: 'cache miss' });
  }
  try {
    Logger.log('[billing] loadPreparedBilling_ raw cache for ' + key + ': ' + cached);
  } catch (err) {
    // ignore logging errors in non-GAS environments
  }
  try {
    const parsed = JSON.parse(cached);
    const rawLength = Array.isArray(parsed && parsed.billingJson) ? parsed.billingJson.length : 0;
    const validation = validatePreparedBillingPayload_(parsed, expectedMonthKey);
    const normalized = normalizePreparedBilling_(parsed);
    const normalizedLength = normalized && Array.isArray(normalized.billingJson) ? normalized.billingJson.length : 0;
    billingLogger_.log('[billing] loadPreparedBilling_ parsed cache summary=' + JSON.stringify({
      cacheKey: key,
      expectedMonthKey,
      parsedMonth: parsed && (parsed.billingMonth || parsed.month),
      normalizedMonth: normalized && normalized.billingMonth,
      rawBillingJsonLength: rawLength,
      normalizedBillingJsonLength: normalizedLength,
      validationReason: validation.reason || null,
      validationOk: validation.ok
    }));
    if (!validation.ok) {
      if (validation.reason === 'billingMonth mismatch' && normalized && expectedMonthKey && Array.isArray(normalized.billingJson)) {
        const corrected = Object.assign({}, normalized, { billingMonth: expectedMonthKey });
        billingLogger_.log('[billing] loadPreparedBilling_ auto-correcting billingMonth mismatch for cache key=' + key);
        savePreparedBilling_(corrected);
        return wrapResult(corrected, validation);
      }
      if (allowInvalid && normalized && Array.isArray(normalized.billingJson)) {
        try {
          Logger.log('[billing] loadPreparedBilling_: allowing invalid cache for ' + key + ' reason=' + validation.reason);
        } catch (err) {
          // ignore logging errors in non-GAS environments
        }
        return wrapResult(Object.assign({}, normalized, { billingMonth: validation.billingMonth || expectedMonthKey }), validation);
      }
      try {
        Logger.log('[billing] loadPreparedBilling_: invalid cache for ' + key + ' reason=' + validation.reason);
      } catch (err) {
        // ignore logging errors in non-GAS environments
      }
      console.warn('[billing] Prepared cache invalid for ' + key + ': ' + validation.reason);
      clearBillingCache_(key);
      return wrapResult(null, validation);
    }
    return wrapResult(Object.assign({}, normalized || parsed, { billingMonth: validation.billingMonth }), validation);
  } catch (err) {
    try {
      Logger.log('[billing] loadPreparedBilling_: failed to parse cache for ' + key + ' error=' + err);
    } catch (logErr) {
      // ignore logging errors in non-GAS environments
    }
    console.warn('[billing] Failed to parse prepared cache', err);
    clearBillingCache_(key);
    return wrapResult(null, { ok: false, reason: 'parse error' });
  }
}

function savePreparedBilling_(payload) {
  const normalizedPayload = normalizePreparedBilling_(payload);
  const resolvedMonthKey = normalizeBillingMonthKeySafe_(normalizedPayload && normalizedPayload.billingMonth);
  if (!normalizedPayload || !resolvedMonthKey) {
    billingLogger_.log('[billing] savePreparedBilling_ skipped due to invalid payload');
    return;
  }
  const payloadToCache = Object.assign({}, normalizedPayload, { billingMonth: resolvedMonthKey });
  const billingJsonLength = Array.isArray(payloadToCache.billingJson) ? payloadToCache.billingJson.length : 0;
  billingLogger_.log('[billing] savePreparedBilling_ summary=' + JSON.stringify({
    billingMonth: payloadToCache.billingMonth,
    billingJsonLength
  }));
  const key = buildBillingCacheKey_(payloadToCache.billingMonth);
  if (!key) return;
  const cache = getBillingCache_();
  if (!cache) return;
  try {
    cache.put(key, JSON.stringify(payloadToCache), BILLING_CACHE_TTL_SECONDS);
  } catch (err) {
    try {
      if (Logger && typeof Logger.warning === 'function') {
        Logger.warning('[billing] Failed to cache prepared billing: ' + err);
      }
    } catch (logErr) {
      // ignore logging errors in non-GAS environments
    }
    console.warn('[billing] Failed to cache prepared billing', err);
  }
}

function buildPreparedBillingPayload_(billingMonth) {
  const resolvedMonthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const source = getBillingSourceData(resolvedMonthKey || billingMonth);
  const billingMonthKey = normalizeBillingMonthKeySafe_(
    source.billingMonth || (source.month && source.month.key) || resolvedMonthKey
  );
  const billingJson = generateBillingJsonFromSource(source);
  const billingJsonArray = Array.isArray(billingJson) ? billingJson : [];
  const visitCounts = source.treatmentVisitCounts || source.visitCounts || {};
  const visitCountKeys = Object.keys(visitCounts || {});
  const normalizeNumber_ = value => {
    const num = Number(value);
    return Number.isFinite(num) ? num : 0;
  };
  const visitsByPatient = visitCountKeys.reduce((map, pid) => {
    const entry = visitCounts[pid];
    const visitCount = entry && entry.visitCount != null ? entry.visitCount : entry;
    map[pid] = billingNormalizeVisitCount_(visitCount);
    return map;
  }, {});
  const totalsByPatient = {};
  billingJsonArray.forEach(item => {
    const pid = billingNormalizePatientId_(item && item.patientId);
    if (!pid) return;
    totalsByPatient[pid] = {
      visitCount: billingNormalizeVisitCount_(item && item.visitCount),
      billingAmount: normalizeNumber_(item && item.billingAmount),
      total: normalizeNumber_(item && item.total),
      grandTotal: normalizeNumber_(item && item.grandTotal),
      carryOverAmount: normalizeNumber_(item && item.carryOverAmount),
      carryOverFromHistory: normalizeNumber_(item && item.carryOverFromHistory)
    };
  });
  const patientMap = source.patients || source.patientMap || {};
  const bankAccountInfoByPatient = Object.keys(patientMap || {}).reduce((map, pid) => {
    const patient = patientMap[pid] || {};
    map[pid] = {
      bankCode: patient.bankCode || '',
      branchCode: patient.branchCode || '',
      accountNumber: patient.accountNumber || '',
      regulationCode: normalizeNumber_(patient.regulationCode || (patient.raw && patient.raw.regulationCode)),
      nameKanji: patient.nameKanji || '',
      nameKana: patient.nameKana || '',
      isNew: patient.isNew != null ? patient.isNew : normalizeNumber_(patient.raw && patient.raw.isNew)
    };
    return map;
  }, {});
  const zeroVisitSamples = visitCountKeys
    .filter(pid => {
      const entry = visitCounts[pid];
      const visitCount = entry && entry.visitCount != null ? entry.visitCount : entry;
      return !visitCount || Number(visitCount) === 0;
    })
    .slice(0, 5);

  billingLogger_.log('[billing] buildPreparedBillingPayload_ summary=' + JSON.stringify({
    billingMonth: billingMonthKey || source.billingMonth,
    treatmentVisitCountEntries: visitCountKeys.length,
    zeroVisitSamples,
    billingJsonLength: billingJsonArray.length
  }));
  if (billingJsonArray.length) {
    billingLogger_.log('[billing] buildPreparedBillingPayload_ firstBillingEntry=' + JSON.stringify(billingJsonArray[0]));
  }
  return {
    schemaVersion: PREPARED_BILLING_SCHEMA_VERSION,
    billingMonth: billingMonthKey || source.billingMonth,
    billingJson: billingJsonArray,
    preparedAt: new Date().toISOString(),
    patients: source.patients || source.patientMap || {},
    bankInfoByName: source.bankInfoByName || {},
    staffByPatient: source.staffByPatient || {},
    staffDirectory: source.staffDirectory || {},
    staffDisplayByPatient: source.staffDisplayByPatient || {},
    billingOverrideFlags: source.billingOverrideFlags || {},
      carryOverByPatient: source.carryOverByPatient || {},
      carryOverLedger: source.carryOverLedger || [],
      carryOverLedgerMeta: source.carryOverLedgerMeta || {},
      carryOverLedgerByPatient: source.carryOverLedgerByPatient || {},
      unpaidHistory: source.unpaidHistory || [],
      visitsByPatient,
      totalsByPatient,
      bankAccountInfoByPatient
  };
}

function coerceBillingJsonArray_(raw) {
  if (Array.isArray(raw)) return raw;
  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
      console.warn('[billing] Failed to parse billingJson string', err);
    }
  }
  return [];
}

  function normalizePreparedBilling_(payload) {
    const rawLength = Array.isArray(payload && payload.billingJson) ? payload.billingJson.length : 0;
    if (!payload) return null;
    const normalizeMap_ = value => (value && typeof value === 'object' && !Array.isArray(value) ? value : {});
    const schemaVersion = Number(payload.schemaVersion);
    const normalized = Object.assign({}, payload, {
      schemaVersion: Number.isFinite(schemaVersion) ? schemaVersion : null,
      billingMonth: payload.billingMonth || payload.month || '',
      preparedAt: payload.preparedAt || payload.prepared_at || null,
      billingJson: coerceBillingJsonArray_(payload.billingJson),
      patients: normalizeMap_(payload.patients || payload.patientMap),
      patientMap: normalizeMap_(payload.patientMap || payload.patients),
      bankInfoByName: normalizeMap_(payload.bankInfoByName),
      bankAccountInfoByPatient: normalizeMap_(payload.bankAccountInfoByPatient),
      visitsByPatient: normalizeMap_(payload.visitsByPatient || payload.visitCounts),
      totalsByPatient: normalizeMap_(payload.totalsByPatient),
      staffByPatient: normalizeMap_(payload.staffByPatient),
      staffDirectory: normalizeMap_(payload.staffDirectory),
      staffDisplayByPatient: normalizeMap_(payload.staffDisplayByPatient),
      billingOverrideFlags: normalizeMap_(payload.billingOverrideFlags),
      carryOverByPatient: normalizeMap_(payload.carryOverByPatient),
      carryOverLedger: Array.isArray(payload.carryOverLedger) ? payload.carryOverLedger : [],
      carryOverLedgerMeta: normalizeMap_(payload.carryOverLedgerMeta),
      carryOverLedgerByPatient: normalizeMap_(payload.carryOverLedgerByPatient),
      unpaidHistory: Array.isArray(payload.unpaidHistory) ? payload.unpaidHistory : [],
      bankStatuses: normalizeMap_(payload.bankStatuses)
    });
    const normalizedLength = Array.isArray(normalized.billingJson) ? normalized.billingJson.length : 0;
    billingLogger_.log('[billing] normalizePreparedBilling_ lengths=' + JSON.stringify({
      rawBillingJsonLength: rawLength,
      normalizedBillingJsonLength: normalizedLength,
      billingMonth: normalized.billingMonth
    }));
    return normalized;
  }

function toClientBillingPayload_(prepared) {
  const rawLength = prepared && prepared.billingJson
    ? (Array.isArray(prepared.billingJson) ? prepared.billingJson.length : 'non-array')
    : 0;
  const normalized = normalizePreparedBilling_(prepared);
  if (!normalized) return null;
  const billingJson = Array.isArray(normalized.billingJson) ? normalized.billingJson : [];
  billingLogger_.log('[billing] toClientBillingPayload_: billingJson length before normalize=' + rawLength +
    ' after normalize=' + billingJson.length);
  return {
    schemaVersion: normalized.schemaVersion,
    billingMonth: normalized.billingMonth || '',
    billingJson,
    preparedAt: normalized.preparedAt || null,
    patients: normalized.patients || {},
    bankInfoByName: normalized.bankInfoByName || {},
    staffByPatient: normalized.staffByPatient || {},
    staffDirectory: normalized.staffDirectory || {},
    staffDisplayByPatient: normalized.staffDisplayByPatient || {},
    billingOverrideFlags: normalized.billingOverrideFlags || {},
      carryOverByPatient: normalized.carryOverByPatient || {},
      carryOverLedger: normalized.carryOverLedger || [],
      carryOverLedgerMeta: normalized.carryOverLedgerMeta || {},
      carryOverLedgerByPatient: normalized.carryOverLedgerByPatient || {},
      unpaidHistory: normalized.unpaidHistory || [],
      visitsByPatient: normalized.visitsByPatient || {},
      totalsByPatient: normalized.totalsByPatient || {},
      bankAccountInfoByPatient: normalized.bankAccountInfoByPatient || {},
    bankStatuses: normalized.bankStatuses || {}
  };
}

function serializeBillingPayload_(payload) {
  if (!payload) return null;
  const beforeLength = payload && payload.billingJson
    ? (Array.isArray(payload.billingJson) ? payload.billingJson.length : 'non-array')
    : 0;
  try {
    const json = JSON.stringify(payload);
    const parsed = JSON.parse(json);
    const afterLength = parsed && parsed.billingJson
      ? (Array.isArray(parsed.billingJson) ? parsed.billingJson.length : 'non-array')
      : 0;
    billingLogger_.log('[billing] serializeBillingPayload_: billingJson length before=' + beforeLength + ' after=' + afterLength);
    return parsed;
  } catch (err) {
    console.warn('[billing] Failed to serialize billing payload', err);
    return null;
  }
}

function prepareBillingData(billingMonth) {
  const normalizedMonth = normalizeBillingMonthInput(billingMonth);
  const prepared = buildPreparedBillingPayload_(normalizedMonth);
  const clientPayload = toClientBillingPayload_(prepared);
  const payloadWithMonth = clientPayload
    ? Object.assign({}, clientPayload, { billingMonth: normalizedMonth.key })
    : clientPayload;
  const serialized = serializeBillingPayload_(payloadWithMonth) || payloadWithMonth;
  const payloadJson = serialized ? JSON.stringify(serialized) : '';
  const payloadPreview = payloadJson.length > 50000 ? payloadJson.slice(0, 50000) + '…<truncated>' : payloadJson;

  billingLogger_.log(
    '[billing] prepareBillingData payloadSummary=' +
      JSON.stringify({
        billingMonth: serialized && serialized.billingMonth,
        preparedAt: serialized && serialized.preparedAt,
        billingJsonLength: serialized && serialized.billingJson ? serialized.billingJson.length : null,
        patientCount: serialized && serialized.patients ? Object.keys(serialized.patients).length : null,
        payloadByteLength: payloadJson.length
      }) +
      '\n[billing] prepareBillingData payloadRaw=' + payloadPreview
  );

  const cachePayload = serialized
    ? Object.assign({}, serialized, { billingMonth: normalizedMonth.key })
    : { billingMonth: normalizedMonth.key };
  savePreparedBilling_(cachePayload);
  savePreparedBillingToSheet_(normalizedMonth.key, cachePayload);
  return cachePayload;
}

/**
 * Generate billing JSON without creating files (preview use case).
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @return {Object} billingMonth key and generated JSON array.
 */
function generateBillingJsonPreview(billingMonth) {
  const prepared = buildPreparedBillingPayload_(billingMonth);
  savePreparedBilling_(prepared);
  savePreparedBillingToSheet_(billingMonth, prepared);
  return prepared;
}

/**
 * Generate invoice PDFs and surface JSON for UI.
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @param {Object} [options] - Optional output overrides such as fileName or note.
 * @return {Object} billingMonth key, billingJson, and saved file metadata.
 */
function generateInvoices(billingMonth, options) {
  try {
    const prepared = buildPreparedBillingPayload_(billingMonth);
    savePreparedBilling_(prepared);
    savePreparedBillingToSheet_(billingMonth, prepared);
    return generatePreparedInvoices_(prepared, options);
  } catch (err) {
    const msg = err && err.message ? err.message : String(err);
    const stack = err && err.stack ? '\n' + err.stack : '';
    console.error('[generateInvoices] failed:', msg, stack);
    throw new Error('請求生成中にエラーが発生しました: ' + msg);
  }
}

function generatePreparedInvoices_(prepared, options) {
  const normalized = normalizePreparedBilling_(prepared);
  if (!normalized || !normalized.billingJson) {
    throw new Error('請求集計結果が見つかりません。先に集計を実行してください。');
  }
  const outputOptions = Object.assign({}, options, { billingMonth: normalized.billingMonth });
  const pdfs = generateInvoicePdfs(normalized.billingJson, outputOptions);
  const shouldExportBank = !outputOptions || outputOptions.skipBankExport !== true;
  const bankOutput = shouldExportBank ? exportBankTransferDataForPrepared_(normalized) : null;
  return {
    billingMonth: normalized.billingMonth,
    billingJson: normalized.billingJson,
    files: pdfs.files,
    bankOutput,
    preparedAt: normalized.preparedAt || null
  };
}

function generateInvoicesFromCache(billingMonth, options) {
  const month = normalizeBillingMonthInput(billingMonth);
  const prepared = normalizePreparedBilling_(loadPreparedBilling_(month.key));
  if (!prepared || !prepared.billingJson) {
    throw new Error('事前集計が見つかりません。先に「請求データを集計」ボタンを実行してください。');
  }
  return generatePreparedInvoices_(prepared, options);
}

function normalizeBillingEditBurden_(value) {
  if (value === null || value === undefined) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  if (raw === '自費') return '自費';
  const num = Number(raw);
  if (Number.isFinite(num) && num > 0) {
    if (num > 0 && num <= 1) return Math.round(num * 10);
    if (num < 10) return Math.round(num);
  }
  return null;
}

function normalizeBillingEditMedicalAssistance_(value) {
  if (value === undefined) return undefined;
  if (value === null) return 0;
  if (value === true) return 1;
  if (value === false) return 0;
  const num = Number(value);
  if (Number.isFinite(num)) return num ? 1 : 0;
  const text = String(value || '').trim().toLowerCase();
  if (!text) return 0;
  return ['1', 'true', 'yes', 'y', 'on', '有', 'あり', '〇', '○', '◯'].indexOf(text) >= 0 ? 1 : 0;
}

function normalizeBillingEdits_(maybeEdits) {
  if (!Array.isArray(maybeEdits)) return [];
  return maybeEdits.map(edit => {
    const pid = edit && edit.patientId ? String(edit.patientId).trim() : '';
    if (!pid) return null;
    const burden = normalizeBillingEditBurden_(edit.burdenRate);
    const bankInfo = edit && typeof edit.bankInfo === 'object' ? edit.bankInfo : edit;
    const normalizedBankCode = bankInfo && bankInfo.bankCode != null ? String(bankInfo.bankCode).trim() : undefined;
    const normalizedBranchCode = bankInfo && bankInfo.branchCode != null ? String(bankInfo.branchCode).trim() : undefined;
    const normalizedAccountNumber = bankInfo && bankInfo.accountNumber != null
      ? String(bankInfo.accountNumber).trim()
      : undefined;
    const unitPriceInput = edit && Object.prototype.hasOwnProperty.call(edit, 'unitPrice')
      ? edit.unitPrice
      : edit.manualUnitPrice;
    const hasManualUnitPriceInput = edit
      && (Object.prototype.hasOwnProperty.call(edit, 'unitPrice')
        || Object.prototype.hasOwnProperty.call(edit, 'manualUnitPrice'));
    const hasManualUnitPrice = hasManualUnitPriceInput && unitPriceInput !== null && unitPriceInput !== '';
    const hasManualTransportInput = edit && (Object.prototype.hasOwnProperty.call(edit, 'manualTransportAmount')
      || Object.prototype.hasOwnProperty.call(edit, 'transportAmount'));
    const manualTransportSource = edit && Object.prototype.hasOwnProperty.call(edit, 'manualTransportAmount')
      ? edit.manualTransportAmount
      : edit.transportAmount;
    const normalizedManualTransport = manualTransportSource === '' || manualTransportSource === null
      ? ''
      : Number(manualTransportSource) || 0;
    const hasManualSelfPayInput = edit && Object.prototype.hasOwnProperty.call(edit, 'manualSelfPayAmount');
    const manualSelfPaySource = hasManualSelfPayInput ? edit.manualSelfPayAmount : edit && edit.selfPayAmount;
    const normalizedManualSelfPay = manualSelfPaySource === '' || manualSelfPaySource === null
      ? ''
      : Number(manualSelfPaySource) || 0;
    const hasAdjustedVisitInput = edit && Object.prototype.hasOwnProperty.call(edit, 'adjustedVisitCount');
    const normalizedAdjustedVisit = hasAdjustedVisitInput
      ? billingNormalizeVisitCount_(edit.adjustedVisitCount)
      : edit && Object.prototype.hasOwnProperty.call(edit, 'visitCount')
        ? billingNormalizeVisitCount_(edit.visitCount)
        : undefined;
    return {
      patientId: pid,
      insuranceType: edit.insuranceType != null ? String(edit.insuranceType).trim() : undefined,
      medicalAssistance: normalizeBillingEditMedicalAssistance_(edit.medicalAssistance),
      burdenRate: burden !== null ? burden : undefined,
      // normalized unitPrice retains legacy behavior for immediate billing recalculation,
      // while manualUnitPrice preserves blank input for persistence.
      unitPrice: hasManualUnitPrice ? Number(unitPriceInput) || 0 : undefined,
      manualUnitPrice: hasManualUnitPriceInput ? unitPriceInput : undefined,
      carryOverAmount: edit.carryOverAmount != null ? Number(edit.carryOverAmount) || 0 : undefined,
      payerType: edit.payerType != null ? String(edit.payerType).trim() : undefined,
      manualTransportAmount: hasManualTransportInput ? normalizedManualTransport : undefined,
      manualSelfPayAmount: hasManualSelfPayInput ? normalizedManualSelfPay : undefined,
      responsible: edit && edit.responsible != null ? String(edit.responsible).trim() : undefined,
      bankCode: normalizedBankCode,
      branchCode: normalizedBranchCode,
      accountNumber: normalizedAccountNumber,
      isNew: edit && Object.prototype.hasOwnProperty.call(edit, 'isNew')
        ? normalizeZeroOneFlag_(edit.isNew)
        : undefined,
      adjustedVisitCount: normalizedAdjustedVisit
    };
  }).filter(Boolean);
}

function splitBillingEditsByTarget_(edits) {
  const normalized = normalizeBillingEdits_(Array.isArray(edits) ? edits : []);
  return {
    patientEdits: normalized.filter(edit => hasDefinedField_(extractPatientInfoUpdateFields_(edit))),
    overrideEdits: normalized.filter(hasBillingOverrideValues_)
  };
}

function normalizeBillingEditsPayload_(options) {
  const opts = options || {};
  if (Array.isArray(opts.patientInfoUpdates) || Array.isArray(opts.billingOverridesUpdates)) {
    const patientEdits = splitBillingEditsByTarget_(opts.patientInfoUpdates || []);
    const overrideEdits = splitBillingEditsByTarget_(opts.billingOverridesUpdates || []);
    return {
      patientEdits: patientEdits.patientEdits,
      overrideEdits: overrideEdits.overrideEdits
    };
  }
  return splitBillingEditsByTarget_(opts.edits);
}

function extractPatientInfoUpdateFields_(edit) {
  return {
    insuranceType: edit.insuranceType,
    burdenRate: edit.burdenRate,
    medicalAssistance: edit.medicalAssistance,
    medicalSubsidy: edit.medicalSubsidy,
    payerType: edit.payerType,
    responsible: edit.responsible,
    bankCode: edit.bankCode,
    branchCode: edit.branchCode,
    accountNumber: edit.accountNumber,
    isNew: edit.isNew
  };
}

function extractBillingOverrideFields_(edit) {
  return {
    patientId: edit.patientId,
    manualUnitPrice: edit.manualUnitPrice,
    manualTransportAmount: edit.manualTransportAmount,
    carryOverAmount: edit.carryOverAmount,
    adjustedVisitCount: edit.adjustedVisitCount,
    manualSelfPayAmount: edit.manualSelfPayAmount
  };
}

function hasBillingOverrideValues_(edit) {
  const fields = extractBillingOverrideFields_(edit);
  return ['manualUnitPrice', 'manualTransportAmount', 'carryOverAmount', 'adjustedVisitCount', 'manualSelfPayAmount']
    .some(key => fields[key] !== undefined);
}

function hasDefinedField_(obj) {
  return Object.keys(obj || {}).some(key => obj[key] !== undefined);
}

function savePatientUpdate(patientId, updatedFields) {
  const pid = billingNormalizePatientId_(patientId);
  if (!pid) return { updated: false };
  const fields = updatedFields || {};

  const sheet = billingSs().getSheetByName(BILLING_PATIENT_SHEET_NAME);
  if (!sheet) {
    throw new Error('患者情報シートが見つかりません: ' + BILLING_PATIENT_SHEET_NAME);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { updated: false };

  const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const colPid = resolveBillingColumn_(headers, BILLING_LABELS.recNo, '患者ID', { required: true, fallbackIndex: BILLING_PATIENT_COLS_FIXED.recNo });
  const colInsurance = resolveBillingColumn_(headers, ['保険区分', '保険種別', '保険タイプ', '保険'], '保険区分', {});
  const colBurden = resolveBillingColumn_(headers, BILLING_LABELS.share, '負担割合', { fallbackIndex: BILLING_PATIENT_COLS_FIXED.share });
  const colUnitPrice = resolveBillingColumn_(headers, ['単価', '請求単価', '自費単価', '単価(自費)', '単価（自費）', '単価（手動上書き）', '単価(手動上書き)'], '単価', {});
  const colCarryOver = resolveBillingColumn_(headers, ['未入金', '未入金額', '未収金', '未収', '繰越', '繰越額', '繰り越し', '差引繰越', '前回未払', '前回未収', 'carryOverAmount'], '未入金額', {});
  const colPayer = resolveBillingColumn_(headers, ['保険者', '支払区分', '保険/自費', '保険区分種別'], '保険者', {});
  const colMedicalSubsidy = resolveBillingColumn_(headers, ['医療助成'], '医療助成', { fallbackLetter: 'AS' });
  const colTransport = resolveBillingColumn_(headers, ['交通費', '交通費(手動)', '交通費（手動）', 'transportAmount', 'manualTransportAmount'], '交通費', {});
  const colBank = resolveBillingColumn_(headers, ['銀行コード', '銀行CD', '銀行番号', 'bankCode'], '銀行コード', { fallbackLetter: 'N' });
  const colBranch = resolveBillingColumn_(headers, ['支店コード', '支店番号', '支店CD', 'branchCode'], '支店コード', { fallbackLetter: 'O' });
  const colAccount = resolveBillingColumn_(headers, ['口座番号', '口座No', '口座NO', 'accountNumber', '口座'], '口座番号', { fallbackLetter: 'Q' });
  const colIsNew = resolveBillingColumn_(headers, ['新規', '新患', 'isNew', '新規フラグ', '新規区分'], '新規区分', { fallbackLetter: 'U' });
  const colResponsible = resolveBillingColumn_(headers, ['担当者', 'responsible', 'responsibleName', '担当', '担当者名'], '担当者', {});

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let idx = 0; idx < values.length; idx++) {
    const row = values[idx];
    const rowPid = billingNormalizePatientId_(row[colPid - 1]);
    if (rowPid !== pid) continue;

    const newRow = row.slice();
    if (colInsurance && fields.insuranceType !== undefined) newRow[colInsurance - 1] = fields.insuranceType;
    const normalizedMedicalFlag = fields.medicalSubsidy !== undefined
      ? normalizeZeroOneFlag_(fields.medicalSubsidy)
      : (fields.medicalAssistance !== undefined ? normalizeBillingEditMedicalAssistance_(fields.medicalAssistance) : undefined);
    if (colMedicalSubsidy && normalizedMedicalFlag !== undefined) newRow[colMedicalSubsidy - 1] = normalizedMedicalFlag ? 1 : 0;
    if (colBurden && fields.burdenRate !== undefined) newRow[colBurden - 1] = fields.burdenRate;
    if (colUnitPrice && fields.manualUnitPrice !== undefined) {
      const isBlank = fields.manualUnitPrice === '' || fields.manualUnitPrice === null;
      newRow[colUnitPrice - 1] = isBlank ? '' : Number(fields.manualUnitPrice) || 0;
    }
    if (colTransport && fields.manualTransportAmount !== undefined) {
      const isBlankTransport = fields.manualTransportAmount === '' || fields.manualTransportAmount === null;
      newRow[colTransport - 1] = isBlankTransport ? '' : Number(fields.manualTransportAmount) || 0;
    }
    if (colCarryOver && fields.carryOverAmount !== undefined) newRow[colCarryOver - 1] = fields.carryOverAmount;
    if (colPayer && fields.payerType !== undefined) newRow[colPayer - 1] = fields.payerType;
    if (colBank && fields.bankCode !== undefined) newRow[colBank - 1] = String(fields.bankCode || '').trim();
    if (colBranch && fields.branchCode !== undefined) newRow[colBranch - 1] = String(fields.branchCode || '').trim();
    if (colAccount && fields.accountNumber !== undefined) newRow[colAccount - 1] = String(fields.accountNumber || '').trim();
    if (colIsNew && fields.isNew !== undefined) newRow[colIsNew - 1] = normalizeZeroOneFlag_(fields.isNew);
    if (colResponsible && fields.responsible !== undefined) newRow[colResponsible - 1] = String(fields.responsible || '').trim();

    sheet.getRange(idx + 2, 1, 1, newRow.length).setValues([newRow]);
    return { updated: true, rowNumber: idx + 2 };
  }

  return { updated: false };
}

function applyBillingPatientEdits_(edits) {
  if (!Array.isArray(edits) || !edits.length) return { updated: 0 };

  let updatedCount = 0;
  edits.forEach(edit => {
    const patientFields = extractPatientInfoUpdateFields_(edit);
    if (!hasDefinedField_(patientFields)) return;
    const result = savePatientUpdate(edit.patientId, patientFields);
    if (result && result.updated) {
      updatedCount += 1;
    }
  });

  return { updated: updatedCount };
}

function ensureBillingOverridesSheet_() {
  const ss = billingSs();
  let sheet = ss.getSheetByName(BILLING_OVERRIDES_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(BILLING_OVERRIDES_SHEET_NAME);
  sheet.appendRow([
    'ym',
    'patientId',
    'manualUnitPrice',
    'manualTransportAmount',
    'carryOverAmount',
    'adjustedVisitCount',
    'manualSelfPayAmount'
  ]);
  return sheet;
}

function resolveBillingOverridesColumns_(headers) {
  const colYm = resolveBillingColumn_(headers, ['ym', 'billingMonth', 'month'], 'ym', { required: true });
  const colPid = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', 'patientId']), 'patientId', { required: true });
  const colManualUnitPrice = resolveBillingColumn_(headers, ['manualUnitPrice', 'unitPrice', '単価'], 'manualUnitPrice', {});
  const colManualTransport = resolveBillingColumn_(headers, ['manualTransportAmount', 'transportAmount', '交通費'], 'manualTransportAmount', {});
  const colCarryOver = resolveBillingColumn_(headers, ['carryOverAmount', 'carryOver', '未入金額', '繰越'], 'carryOverAmount', {});
  const colAdjustedVisits = resolveBillingColumn_(headers, ['adjustedVisitCount', 'visitCount', '回数'], 'adjustedVisitCount', {});
  const colManualSelfPay = resolveBillingColumn_(headers, ['manualSelfPayAmount', 'selfPayAmount', '自費'], 'manualSelfPayAmount', {});
  return { colYm, colPid, colManualUnitPrice, colManualTransport, colCarryOver, colAdjustedVisits, colManualSelfPay };
}

function saveBillingOverrideUpdate_(billingMonthKey, edit) {
  const ym = billingMonthKey ? String(billingMonthKey).trim() : '';
  const pid = billingNormalizePatientId_(edit && edit.patientId);
  if (!ym || !pid) return { updated: false };
  const sheet = ensureBillingOverridesSheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const cols = resolveBillingOverridesColumns_(headers);
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  const assignField = (row, col, value) => {
    if (!col || value === undefined) return;
    row[col - 1] = value;
  };

  for (let idx = 0; idx < values.length; idx++) {
    const row = values[idx];
    const rowYm = String(row[cols.colYm - 1] || '').trim();
    const rowPid = billingNormalizePatientId_(row[cols.colPid - 1]);
    if (rowYm !== ym || rowPid !== pid) continue;

    const newRow = row.slice();
    assignField(newRow, cols.colYm, ym);
    assignField(newRow, cols.colPid, pid);
    assignField(newRow, cols.colManualUnitPrice, edit.manualUnitPrice);
    assignField(newRow, cols.colManualTransport, edit.manualTransportAmount);
    assignField(newRow, cols.colCarryOver, edit.carryOverAmount);
    assignField(newRow, cols.colAdjustedVisits, edit.adjustedVisitCount);
    assignField(newRow, cols.colManualSelfPay, edit.manualSelfPayAmount);

    sheet.getRange(idx + 2, 1, 1, newRow.length).setValues([newRow]);
    return { updated: true, rowNumber: idx + 2 };
  }

  const newRow = new Array(Math.max(lastCol, 7)).fill('');
  assignField(newRow, cols.colYm, ym);
  assignField(newRow, cols.colPid, pid);
  assignField(newRow, cols.colManualUnitPrice, edit.manualUnitPrice);
  assignField(newRow, cols.colManualTransport, edit.manualTransportAmount);
  assignField(newRow, cols.colCarryOver, edit.carryOverAmount);
  assignField(newRow, cols.colAdjustedVisits, edit.adjustedVisitCount);
  assignField(newRow, cols.colManualSelfPay, edit.manualSelfPayAmount);

  sheet.appendRow(newRow);
  return { updated: true, rowNumber: sheet.getLastRow() };
}

function applyBillingOverrideEdits_(billingMonthKey, edits) {
  if (!Array.isArray(edits) || !edits.length) return { updated: 0 };
  let updatedCount = 0;
  edits.forEach(edit => {
    if (!hasBillingOverrideValues_(edit)) return;
    const normalizedCarryOver = edit.carryOverAmount === '' || edit.carryOverAmount === null
      ? ''
      : edit.carryOverAmount;
    const normalizedManualUnitPrice = edit.manualUnitPrice === '' || edit.manualUnitPrice === null
      ? ''
      : edit.manualUnitPrice;
    const normalizedManualTransport = edit.manualTransportAmount === '' || edit.manualTransportAmount === null
      ? ''
      : edit.manualTransportAmount;
    const normalizedManualSelfPay = edit.manualSelfPayAmount === '' || edit.manualSelfPayAmount === null
      ? ''
      : edit.manualSelfPayAmount;
    const result = saveBillingOverrideUpdate_(billingMonthKey, Object.assign({}, edit, {
      carryOverAmount: normalizedCarryOver,
      manualUnitPrice: normalizedManualUnitPrice,
      manualTransportAmount: normalizedManualTransport,
      manualSelfPayAmount: normalizedManualSelfPay
    }));
    if (result && result.updated) {
      updatedCount += 1;
    }
  });
  return { updated: updatedCount };
}

function applyBillingEdits(billingMonth, options) {
  const opts = options || {};
  const normalized = normalizeBillingEditsPayload_(opts);
  const patientEdits = normalized && normalized.patientEdits ? normalized.patientEdits : [];
  const overrideEdits = normalized && normalized.overrideEdits ? normalized.overrideEdits : [];
  let refreshedPatients = null;
  if (patientEdits.length) {
    applyBillingPatientEdits_(patientEdits);
    refreshedPatients = getBillingPatientRecords();
  }
  if (overrideEdits.length) {
    applyBillingOverrideEdits_(billingMonth, overrideEdits);
  }
  const prepared = buildPreparedBillingPayload_(billingMonth);
  if (refreshedPatients) {
    prepared.patients = indexByPatientId_(refreshedPatients);
  }
  savePreparedBilling_(prepared);
  savePreparedBillingToSheet_(billingMonth, prepared);
  return prepared;
}

function applyBillingEditsAndGenerateInvoices(billingMonth, options) {
  const prepared = applyBillingEdits(billingMonth, options);
  return generatePreparedInvoices_(prepared, options || {});
}

  function generateBankTransferData(billingMonth, options) {
    const prepared = buildPreparedBillingPayload_(billingMonth);
    savePreparedBilling_(prepared);
    savePreparedBillingToSheet_(billingMonth, prepared);
    return exportBankTransferDataForPrepared_(prepared, options || {});
  }

  function generateBankTransferDataFromCache(billingMonth, options) {
    const opts = options || {};
    const monthInput = billingMonth || opts.billingMonth || (opts.prepared && (opts.prepared.billingMonth || opts.prepared.month));
    const resolvedMonthKey = normalizeBillingMonthKeySafe_(monthInput);
    const month = resolvedMonthKey ? normalizeBillingMonthInput(resolvedMonthKey) : null;
    const preparedFromSheet = month ? loadPreparedBillingFromSheet_(month.key) : null;
    const validation = validatePreparedBillingPayload_(preparedFromSheet, month && month.key);
    const normalizedPrepared = normalizePreparedBilling_(preparedFromSheet);

    try {
      billingLogger_.log('[bankExport] generateBankTransferDataFromCache summary=' + JSON.stringify({
        requestedMonth: monthInput || null,
        resolvedMonthKey: month ? month.key : null,
        hasPrepared: !!preparedFromSheet,
        preparedBillingJsonLength: preparedFromSheet && Array.isArray(preparedFromSheet.billingJson)
          ? preparedFromSheet.billingJson.length
          : null,
        validationReason: validation && validation.reason ? validation.reason : null,
        validationOk: validation && validation.hasOwnProperty('ok') ? validation.ok : null
      }));
    } catch (err) {
      // ignore logging errors in non-GAS environments
    }

    if (!month) {
      throw new Error('銀行データを出力できません。請求月が指定されていません。先に請求データを集計・確定してください。');
    }

    if (!validation || !validation.ok || !normalizedPrepared || !Array.isArray(normalizedPrepared.billingJson)) {
      throw new Error(resolveBankExportErrorMessage_(validation));
    }

    if (normalizedPrepared.billingJson.length === 0) {
      return createEmptyBankTransferResult_(month.key, '当月の請求対象はありません');
    }

    const preparedWithMonth = Object.assign({}, normalizedPrepared, {
      billingMonth: month.key
    });

    return exportBankTransferDataForPrepared_(preparedWithMonth, Object.assign({}, opts, {
      billingMonth: month.key
    }));
  }

  function createEmptyBankTransferResult_(billingMonth, message) {
    const resolvedMonth = typeof billingMonth === 'string' ? billingMonth : (billingMonth && billingMonth.key) || '';
    return { billingMonth: resolvedMonth, rows: [], inserted: 0, skipped: 0, message: message || '' };
  }

function resolveBankExportErrorMessage_(validation) {
  const reason = validation && validation.reason ? String(validation.reason) : '';
  const ledgerReasons = ['carryOverLedger missing', 'carryOverLedgerByPatient missing', 'carryOverLedgerMeta missing', 'carryOverByPatient missing', 'unpaidHistory missing'];
  const missingReasons = ['payload missing', 'billingMonth missing', 'cache key missing', 'cache unavailable', 'cache miss', 'parse error'];
  try {
    billingLogger_.log('[bankExport] resolveBankExportErrorMessage_ reason=' + reason);
  } catch (err) {
    // ignore logging errors in non-GAS environments
  }
  if (ledgerReasons.indexOf(reason) >= 0) {
    return '銀行データを生成できません。繰越金データを確認してください。';
  }
  if (reason === 'billingMonth mismatch') {
    return '銀行データを生成できません。請求月が一致する請求データを集計・確定してください。';
  }
  if (reason && missingReasons.indexOf(reason) === -1) {
    return '銀行データを生成できません。請求データは存在しますが、検証に失敗しました。';
  }
  return '銀行データを出力できません。先に請求データを集計・確定してください。';
}

function applyBillingPaymentResultsEntry(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const bankStatuses = getBillingPaymentResults(month);
  return applyPaymentResultsToHistory(month.key, bankStatuses);
}

function applyBillingPaymentResultPdfEntry(billingMonth, fileId, options) {
  const month = normalizeBillingMonthInput(billingMonth);
  const source = getBillingSourceData(month);
  const billingJson = generateBillingJsonFromSource(source);
  const file = DriveApp.getFileById(fileId);
  const pdfBlob = file.getBlob();
  const result = applyPaymentResultPdf(month.key, pdfBlob, billingJson);
  return Object.assign({}, result, {
    fileId,
    fileName: file.getName(),
    billingJson
  });
}

function extractDriveFileId_(input) {
  const raw = String(input || '').trim();
  if (!raw) return '';
  const match = raw.match(/[A-Za-z0-9_-]{25,}/);
  return match ? match[0] : '';
}

function promptBillingResultPdfFileId_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('入金結果PDFのURLまたはファイルIDを入力してください',
    '例: https://drive.google.com/file/d/XXXXX/view', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    throw new Error('入金結果PDF入力がキャンセルされました');
  }
  const fileId = extractDriveFileId_(response.getResponseText());
  if (!fileId) {
    throw new Error('Drive ファイルIDが特定できませんでした');
  }
  return fileId;
}

function promptBillingMonthInput_() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const defaultYm = Utilities.formatDate(today, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMM');
  const response = ui.prompt('請求月を入力してください (YYYYMM)', defaultYm, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    throw new Error('請求月入力がキャンセルされました');
  }
  return normalizeBillingMonthInput(response.getResponseText());
}

function billingGenerateJsonFromMenu() {
  const month = promptBillingMonthInput_();
  const result = generateBillingJsonPreview(month.key);
  SpreadsheetApp.getUi().alert('請求データを生成しました: ' + (result.billingJson ? result.billingJson.length : 0) + ' 件');
  return result;
}

function billingApplyPaymentResultsFromMenu() {
  const month = promptBillingMonthInput_();
  const result = applyBillingPaymentResultsEntry(month.key);
  SpreadsheetApp.getUi().alert('請求履歴を更新しました: ' + result.updated + ' 件');
  return result;
}

function billingApplyPaymentResultPdfFromMenu() {
  const month = promptBillingMonthInput_();
  const fileId = promptBillingResultPdfFileId_();
  const result = applyBillingPaymentResultPdfEntry(month.key, fileId);
  SpreadsheetApp.getUi().alert([
    '入金結果PDFを解析しました。',
    '解析件数: ' + result.parsedCount,
    'マッチ件数: ' + result.matched,
    '履歴更新: ' + result.updated + ' 件'
  ].join('\n'));
  return result;
}

function summarizeBillingHistory_(billingMonth) {
  const sheet = ss().getSheetByName('請求履歴');
  if (!sheet) {
    return { billingMonth, exists: false };
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { billingMonth, exists: true, total: 0, statuses: {}, paidTotal: 0, unpaidTotal: 0 };
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const rows = values.filter(row => row[0] === billingMonth);
  const statuses = {};
  let paidTotal = 0;
  let unpaidTotal = 0;
  rows.forEach(row => {
    const status = row[8] || '';
    statuses[status] = (statuses[status] || 0) + 1;
    paidTotal += Number(row[6]) || 0;
    unpaidTotal += Number(row[7]) || 0;
  });

  return { billingMonth, exists: true, total: rows.length, statuses, paidTotal, unpaidTotal };
}

function billingCheckHistoryFromMenu() {
  const month = promptBillingMonthInput_();
  const summary = summarizeBillingHistory_(month.key);
  const ui = SpreadsheetApp.getUi();
  if (!summary.exists) {
    ui.alert('請求履歴シートが存在しません。');
    return summary;
  }
  const statusTexts = Object.keys(summary.statuses || {}).filter(key => key)
    .map(key => key + ': ' + summary.statuses[key] + ' 件');
  const messageLines = [
    '請求月: ' + summary.billingMonth,
    '履歴件数: ' + summary.total,
    '入金合計: ' + summary.paidTotal.toLocaleString('ja-JP') + ' 円',
    '未入金合計: ' + summary.unpaidTotal.toLocaleString('ja-JP') + ' 円'
  ];
  if (statusTexts.length) {
    messageLines.push('ステータス内訳:');
    messageLines.push(statusTexts.join(', '));
  }
  ui.alert(messageLines.join('\n'));
  return summary;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const billingMenu = ui.createMenu('請求処理');
  billingMenu.addItem('請求データ生成（JSON）', 'billingGenerateJsonFromMenu');
  billingMenu.addItem('入金結果PDFの反映（アップロード→解析）', 'billingApplyPaymentResultPdfFromMenu');
  billingMenu.addItem('（管理者向け）履歴チェック', 'billingCheckHistoryFromMenu');
  billingMenu.addToUi();

  const attendanceMenu = ui.createMenu('勤怠管理');
  attendanceMenu.addItem('勤怠データを今すぐ同期', 'runVisitAttendanceSyncJobFromMenu');
  attendanceMenu.addItem('日次同期トリガーを確認', 'ensureVisitAttendanceSyncTriggerFromMenu');
  attendanceMenu.addToUi();
}
