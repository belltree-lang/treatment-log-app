function doGet(e) {
  return handleBillingDoGet_(e);
}

/**
 * Entry point for the billing web application.
 *
 * Dashboard GAS lives in a separate project and must not be referenced here.
 * Any attempt to route dashboard requests through this endpoint is ignored and
 * logged so that future regressions are visible without breaking billing.
 *
 * @param {Object} e - Apps Script request object.
 * @return {HtmlOutput}
 */
function handleBillingDoGet_(e) {
  const request = e || {};
  guardAgainstDashboardRequest_(request);

  const template = HtmlService.createTemplateFromFile('billing');
  template.baseUrl = ScriptApp.getService().getUrl() || '';
  template.patientId = request.parameter && request.parameter.id ? request.parameter.id : '';
  template.payrollPdfData = {};

  if (request.parameter && request.parameter.lead) template.lead = request.parameter.lead;

  return template
    .evaluate()
    .setTitle('請求処理アプリ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Warn when a dashboard-specific route is accidentally passed to the billing app.
 *
 * This guard is intentionally non-fatal so that historical URLs do not break
 * billing, while still making the separation between Dashboard and Billing
 * explicit in logs.
 */
function guardAgainstDashboardRequest_(e) {
  const path = (e && e.pathInfo ? String(e.pathInfo) : '').replace(/^\/+|\/+$/g, '').toLowerCase();
  const view = e && e.parameter ? String(e.parameter.view || '').toLowerCase() : '';
  const action = e && e.parameter ? String(e.parameter.action || e.parameter.api || '').toLowerCase() : '';

  if (path === 'dashboard' || view === 'dashboard' || action.indexOf('dashboard') !== -1) {
    console.warn('[billing] dashboard routes are not served in this project');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Return the active spreadsheet (utility retained for compatibility).
 * @return {SpreadsheetApp.Spreadsheet}
 */
function ss() {
  try {
    if (typeof SpreadsheetApp === 'undefined' || !SpreadsheetApp.getActiveSpreadsheet) {
      return null;
    }
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    console.warn('[billing] SpreadsheetApp unavailable', err);
    return null;
  }
}

function billingSs() {
  if (typeof resolveBillingSpreadsheet_ === 'function') {
    try {
      const resolved = resolveBillingSpreadsheet_();
      if (resolved) return resolved;
    } catch (err) {
      console.warn('[billing] resolveBillingSpreadsheet_ failed in billingSs', err);
    }
  }

  return ss();
}

function resolveBillingSpreadsheet_() {
  const scriptProps = typeof PropertiesService !== 'undefined'
    ? PropertiesService.getScriptProperties()
    : null;
  const configuredId = (scriptProps && scriptProps.getProperty('SSID'))
    || (typeof APP !== 'undefined' ? (APP.SSID || '') : '');

  if (configuredId) {
    try {
      return SpreadsheetApp.openById(configuredId);
    } catch (err) {
      console.warn('[billing] Failed to open SSID from config:', configuredId, err);
    }
  }

  if (typeof ss === 'function') {
    try {
      const workbook = ss();
      if (workbook) return workbook;
    } catch (err2) {
      console.warn('[billing] Fallback ss() failed', err2);
    }
  }

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
const BILLING_CACHE_CHUNK_MARKER = 'chunked:';
const BILLING_CACHE_MAX_ENTRY_LENGTH = 90000;
const BILLING_CACHE_CHUNK_SIZE = 90000;
const PREPARED_BILLING_SCHEMA_VERSION = 2;
const BANK_INFO_SHEET_NAME = '銀行情報';
const UNPAID_HISTORY_SHEET_NAME = '未回収履歴';
const BANK_WITHDRAWAL_UNPAID_HEADER = '未回収チェック';
const BANK_WITHDRAWAL_AGGREGATE_HEADER = '合算';
const BILLING_DEBUG_PID = '';

if (typeof globalThis !== 'undefined') {
  globalThis.BILLING_CACHE_CHUNK_MARKER = BILLING_CACHE_CHUNK_MARKER;
  globalThis.BILLING_CACHE_MAX_ENTRY_LENGTH = BILLING_CACHE_MAX_ENTRY_LENGTH;
  globalThis.BILLING_CACHE_CHUNK_SIZE = BILLING_CACHE_CHUNK_SIZE;
}

const BILLING_MONTH_KEY_CACHE_ = {};

function shouldLogReceiptDebug_(patientId) {
  const debugPid = String(BILLING_DEBUG_PID || '').trim();
  if (!debugPid) return false;
  const pid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(patientId)
    : String(patientId || '').trim();
  return pid && pid === debugPid;
}

function logReceiptDebug_(patientId, payload) {
  if (!shouldLogReceiptDebug_(patientId)) return;
  const line = '[receipt-debug] ' + JSON.stringify(payload);
  if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
    billingLogger_.log(line);
    return;
  }
  if (typeof console !== 'undefined' && typeof console.log === 'function') {
    console.log(line);
  }
}

function buildBillingMonthKeyCacheKey_(candidate) {
  if (candidate && typeof candidate === 'object') {
    if (candidate.key) return String(candidate.key);
    if (candidate.billingMonth) return String(candidate.billingMonth);
    if (candidate.ym) return String(candidate.ym);
    if (candidate.month && candidate.month.key) return String(candidate.month.key);
    if (candidate.month && candidate.month.ym) return String(candidate.month.ym);
  }
  if (candidate instanceof Date && !isNaN(candidate.getTime())) {
    return 'date:' + candidate.getTime();
  }
  if (typeof candidate === 'string' || typeof candidate === 'number') {
    return String(candidate);
  }
  return '';
}

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
    const cacheKey = buildBillingMonthKeyCacheKey_(candidate);
    if (cacheKey && Object.prototype.hasOwnProperty.call(BILLING_MONTH_KEY_CACHE_, cacheKey)) {
      return BILLING_MONTH_KEY_CACHE_[cacheKey];
    }
    try {
      const normalized = normalizeBillingMonthInput(candidate).key;
      if (cacheKey) {
        BILLING_MONTH_KEY_CACHE_[cacheKey] = normalized;
      }
      return normalized;
    } catch (err) {
      const fallback = String(candidate || '').trim();
      if (fallback) {
        if (cacheKey) {
          BILLING_MONTH_KEY_CACHE_[cacheKey] = fallback;
        }
        return fallback;
      }
    }
  }
  return '';
}

function buildBillingCacheKey_(billingMonthKey) {
  const monthKey = String(billingMonthKey || '').trim();
  if (!monthKey) return '';
  return BILLING_CACHE_PREFIX + monthKey;
}

function buildBillingCacheChunkKey_(baseKey, chunkIndex) {
  return baseKey + ':chunk:' + chunkIndex;
}

function getBillingCache_() {
  try {
    return CacheService.getScriptCache();
  } catch (err) {
    console.warn('[billing] CacheService unavailable', err);
    return null;
  }
}

function isBillingDebugEnabled_() {
  try {
    if (typeof getConfig === 'function') {
      const raw = getConfig('BILLING_DEBUG') || getConfig('billing_debug') || getConfig('BILLING_DEBUG_LOG');
      return String(raw || '').trim() === '1';
    }
  } catch (err) {
    // ignore property access errors
  }
  return false;
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
  let totalPatientCount = 0;
  let totalEmails = 0;
  let resolvedNames = 0;
  Object.keys(staffByPatient || {}).forEach(pid => {
    totalPatientCount += 1;
    const emails = Array.isArray(staffByPatient[pid]) ? staffByPatient[pid] : [staffByPatient[pid]];
    const seen = new Set();
    const names = [];
    emails.forEach(email => {
      const key = billingNormalizeStaffKey_(email);
      if (!key || seen.has(key)) return;
      seen.add(key);
      const resolved = directory[key] || '';
      names.push(resolved || email || '');
      totalEmails += 1;
      if (resolved) resolvedNames += 1;

    });
    result[pid] = names.filter(Boolean);
  });
  billingLogger_.log('[billing] buildStaffDisplayByPatient_: summary=' + JSON.stringify({
    patientCount: totalPatientCount,
    staffEntries: totalEmails,
    resolvedNames
  }));
  return result;
}

function clearBillingCache_(key) {
  if (!key) return;
  const cache = getBillingCache_();
  if (!cache) return;
  try {
    const keysToRemove = [key];
    if (typeof cache.get === 'function') {
      const cached = cache.get(key);
      const chunkCount = parseBillingCacheChunkCount_(cached);
      for (let idx = 1; idx <= chunkCount; idx++) {
        keysToRemove.push(buildBillingCacheChunkKey_(key, idx));
      }
    }
    if (typeof cache.removeAll === 'function') {
      cache.removeAll(keysToRemove);
    } else if (typeof cache.remove === 'function') {
      keysToRemove.forEach(k => cache.remove(k));
    }
  } catch (err) {
    console.warn('[billing] Failed to clear prepared cache', err);
  }
}

function ensurePreparedBillingMetaSheet_() {
  const SHEET_NAME = 'PreparedBillingMeta';
  const HEADER = ['billingMonth', 'preparedAt', 'preparedBy', 'payloadVersion', 'note'];
  const workbook = billingSs();
  if (!workbook) return null;
  let sheet = workbook.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }
  return sheet;
}

function ensurePreparedBillingMetaJsonSheet_() {
  const SHEET_NAME = 'PreparedBillingMetaJson';
  const HEADER = ['billingMonth', 'chunkIndex', 'payloadChunk'];
  const workbook = billingSs();
  if (!workbook) return null;
  let sheet = workbook.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }
  return sheet;
}

function ensurePreparedBillingJsonSheet_() {
  const SHEET_NAME = 'PreparedBillingJson';
  const HEADER = ['billingMonth', 'patientId', 'billingRowJson'];
  const workbook = billingSs();
  if (!workbook) return null;
  let sheet = workbook.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }
  return sheet;
}

function savePreparedBillingMeta_(billingMonth, meta) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) return null;

  const sheet = ensurePreparedBillingMetaSheet_();
  if (!sheet) return null;
  const preparedAtValue = (() => {
    const parsed = meta && meta.preparedAt ? new Date(meta.preparedAt) : null;
    return parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : new Date();
  })();
  const preparedBy = (() => {
    if (meta && meta.preparedBy) return meta.preparedBy;
    try {
      const active = Session.getActiveUser && Session.getActiveUser();
      return active && typeof active.getEmail === 'function' ? active.getEmail() : '';
    } catch (err) {
      return '';
    }
  })();
  const payloadVersion = meta && meta.payloadVersion ? meta.payloadVersion : PREPARED_BILLING_SCHEMA_VERSION;
  const note = meta && meta.note ? meta.note : '';
  const rowValues = [monthKey, preparedAtValue, preparedBy, payloadVersion, note];

  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(row => String(row[0] || '').trim())
    : [];
  const existingIndex = existing.findIndex(value => value === monthKey);

  if (existingIndex >= 0) {
    const targetRow = existingIndex + 2;
    sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    savePreparedBillingMetaJson_(monthKey, meta && meta.metaPayload ? meta.metaPayload : null);
    return { billingMonth: monthKey, row: targetRow, updated: true };
  }

  sheet.insertRows(2, 1);
  sheet.getRange(2, 1, 1, rowValues.length).setValues([rowValues]);
  savePreparedBillingMetaJson_(monthKey, meta && meta.metaPayload ? meta.metaPayload : null);
  return { billingMonth: monthKey, row: 2, updated: false };
}

function savePreparedBillingMetaJson_(billingMonth, metaPayload) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const sheet = ensurePreparedBillingMetaJsonSheet_();
  if (!monthKey) return { billingMonth: monthKey || '', inserted: 0 };
  if (!sheet) return { billingMonth: monthKey, inserted: 0 };

  const lastRow = sheet.getLastRow();
  const existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];

  for (let i = existingRows.length - 1; i >= 0; i--) {
    if (String(existingRows[i][0] || '').trim() === monthKey) {
      sheet.deleteRow(i + 2);
    }
  }

  if (!metaPayload || typeof metaPayload !== 'object') {
    return { billingMonth: monthKey, inserted: 0 };
  }

  const payloadJson = JSON.stringify(metaPayload);
  const chunkSize = 40000;
  const chunks = [];
  for (let i = 0; i < payloadJson.length; i += chunkSize) {
    chunks.push(payloadJson.slice(i, i + chunkSize));
  }

  if (!chunks.length) return { billingMonth: monthKey, inserted: 0 };

  sheet.insertRows(2, chunks.length);
  const rows = chunks.map((chunk, idx) => [monthKey, idx + 1, chunk]);
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  return { billingMonth: monthKey, inserted: rows.length };
}

function getPreparedBillingMonths() {
  // preparedMonths は PreparedBillingMeta シートを取得元とする。
  const sheet = ensurePreparedBillingMetaSheet_();
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const uniqueMonths = new Set(
    values
      .map(row => normalizeBillingMonthKeySafe_(row[0]))
      .filter(Boolean)
  );
  return Array.from(uniqueMonths).sort((a, b) => Number(b) - Number(a));
}

function normalizeReceiptStatus_(value) {
  const status = value == null ? '' : String(value).trim().toUpperCase();
  const allowed = ['UNPAID', 'AGGREGATE', 'HOLD'];
  return allowed.indexOf(status) >= 0 ? status : '';
}

function mergeReceiptSettingsIntoPrepared_(prepared, status, aggregateUntilMonth) {
  const normalizedStatus = normalizeReceiptStatus_(status);
  const normalizedAggregate = normalizedStatus === 'AGGREGATE'
    ? normalizeBillingMonthKeySafe_(aggregateUntilMonth || prepared && prepared.aggregateUntilMonth)
    : '';

  const payload = Object.assign({}, prepared, {
    receiptStatus: normalizedStatus,
    aggregateUntilMonth: normalizedAggregate
  });

  if (Array.isArray(prepared && prepared.billingJson)) {
    payload.billingJson = prepared.billingJson.map(row => Object.assign({}, row || {}, {
      receiptStatus: normalizedStatus,
      aggregateUntilMonth: normalizedAggregate
    }));
  }

  return payload;
  }

  function resolvePreviousBillingMonthKey_(billingMonth) {
    const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
    if (!monthKey) return '';

    const yearNum = Number(monthKey.slice(0, 4));
    const monthNum = Number(monthKey.slice(4, 6));
    if (!Number.isFinite(yearNum) || !Number.isFinite(monthNum)) return '';

    const prevDate = new Date(yearNum, monthNum - 2, 1);
    const prevYear = String(prevDate.getFullYear()).padStart(4, '0');
    const monthText = String(prevDate.getMonth() + 1).padStart(2, '0');
    return prevYear + monthText;
  }

function normalizeReceiptMonthKeys_(months) {
  if (!Array.isArray(months)) return [];
  const seen = new Set();
  const normalized = [];
  months.forEach(month => {
    const normalizedMonth = normalizeBillingMonthKeySafe_(month);
    if (!normalizedMonth || seen.has(normalizedMonth)) return;
    seen.add(normalizedMonth);
    normalized.push(normalizedMonth);
  });
  return normalized;
}

function normalizePastBillingMonths_(months, billingMonth) {
  const monthList = Array.isArray(months) ? months : [];
  const billingKey = normalizeBillingMonthKeySafe_(billingMonth);
  const billingNum = Number(billingKey) || 0;
  const isAggregate = monthList.length > 1;
  const seen = new Set();
  const normalized = [];

  monthList.forEach(value => {
    const ym = normalizeBillingMonthKeySafe_(value);
    if (!ym || seen.has(ym)) return;
    const ymNum = Number(ym) || 0;
    if (billingNum) {
      if (ymNum > billingNum) return;
      if (!isAggregate && ymNum === billingNum) return;
    }
    seen.add(ym);
    normalized.push(ym);
  });

  return normalized.sort();
}

function createBillingMonthCache_() {
  return {
    preparedByMonth: {},
    preparedEntriesByMonth: {},
    bankWithdrawalUnpaidByMonth: {},
    bankWithdrawalAmountsByMonth: {}
  };
}

function getPreparedBillingForMonthCached_(billingMonth, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) return null;
  const store = cache && cache.preparedByMonth ? cache.preparedByMonth : null;
  if (store && Object.prototype.hasOwnProperty.call(store, monthKey)) {
    return store[monthKey];
  }

  const cachedPrepared = loadPreparedBilling_(monthKey, { withValidation: false, allowInvalid: true });
  const cachedPayload = cachedPrepared && cachedPrepared.prepared !== undefined ? cachedPrepared.prepared : cachedPrepared;
  const normalized = cachedPayload ? normalizePreparedBilling_(cachedPayload) : null;
  let summary = normalized || loadPreparedBillingSummaryFromSheet_(monthKey);
  if (!summary && typeof loadPreparedBillingWithSheetFallback_ === 'function') {
    const fallback = loadPreparedBillingWithSheetFallback_(monthKey, { allowInvalid: true, restoreCache: false });
    const fallbackPayload = fallback && fallback.prepared !== undefined ? fallback.prepared : fallback;
    summary = fallbackPayload ? normalizePreparedBilling_(fallbackPayload) : summary;
  }
  const reduced = reducePreparedBillingSummary_(summary);
  if (store) {
    store[monthKey] = reduced || null;
  }
  return reduced;
}

function getPreparedBillingEntryForMonthCached_(billingMonth, patientId, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = billingNormalizePatientId_(patientId);
  if (!monthKey || !pid) return null;
  const store = cache && cache.preparedEntriesByMonth ? cache.preparedEntriesByMonth : null;
  const monthStore = store && store[monthKey];
  if (monthStore && Object.prototype.hasOwnProperty.call(monthStore, pid)) {
    return monthStore[pid] || null;
  }
  const entryMap = loadPreparedBillingEntryMapFromSheet_(monthKey);
  if (store) {
    store[monthKey] = entryMap || {};
  }
  return entryMap && entryMap[pid] ? entryMap[pid] : null;
}

function reducePreparedBillingSummary_(payload) {
  if (!payload || typeof payload !== 'object') return null;
  let totalsByPatient = payload.totalsByPatient || {};
  if (!Object.keys(totalsByPatient || {}).length && Array.isArray(payload.billingJson)) {
    totalsByPatient = {};
    payload.billingJson.forEach(entry => {
      const pid = billingNormalizePatientId_(entry && entry.patientId);
      if (!pid) return;
      totalsByPatient[pid] = pickPreparedBillingEntrySummary_(entry);
    });
  }
  return {
    billingMonth: payload.billingMonth || '',
    preparedAt: payload.preparedAt || null,
    preparedBy: payload.preparedBy || '',
    schemaVersion: payload.schemaVersion || PREPARED_BILLING_SCHEMA_VERSION,
    bankFlagsByPatient: payload.bankFlagsByPatient || {},
    totalsByPatient,
    aggregateUntilMonth: normalizeBillingMonthKeySafe_(payload.aggregateUntilMonth)
  };
}

function loadPreparedBillingSummaryFromSheet_(billingMonth) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const workbook = billingSs();
  if (!workbook) return null;
  const metaSheet = workbook.getSheetByName('PreparedBillingMeta');
  if (!monthKey || !metaSheet) return null;

  const metaLastRow = metaSheet.getLastRow();
  if (metaLastRow < 2) return null;

  const metaValues = metaSheet.getRange(2, 1, metaLastRow - 1, 5).getValues();
  const metaRow = metaValues.find(row => String(row[0] || '').trim() === monthKey);
  if (!metaRow) return null;

  const preparedAtCell = metaRow[1];
  const preparedByCell = metaRow[2];
  const payloadVersion = metaRow[3];
  let parsed = {};
  const metaJsonSheet = workbook.getSheetByName('PreparedBillingMetaJson');
  if (metaJsonSheet) {
    const metaLast = metaJsonSheet.getLastRow();
    if (metaLast >= 2) {
      const rows = metaJsonSheet.getRange(2, 1, metaLast - 1, 3).getValues();
      const chunks = rows
        .filter(row => String(row[0] || '').trim() === monthKey)
        .sort((a, b) => Number(a[1]) - Number(b[1]))
        .map(row => row[2] || '')
        .filter(Boolean);
      const payloadJson = chunks.join('');
      if (payloadJson) {
        try {
          parsed = JSON.parse(payloadJson) || {};
        } catch (err) {
          billingLogger_.log('[billing] loadPreparedBillingSummaryFromSheet_ failed to parse meta payload for ' + monthKey + ': ' + err);
        }
      }
    }
  }

  if (!parsed || typeof parsed !== 'object') {
    parsed = {};
  }

  return Object.assign({}, parsed, {
    billingMonth: monthKey,
    preparedAt: parsed.preparedAt || (preparedAtCell instanceof Date ? preparedAtCell.toISOString() : null),
    preparedBy: parsed.preparedBy || preparedByCell || '',
    schemaVersion: parsed.schemaVersion || payloadVersion || PREPARED_BILLING_SCHEMA_VERSION
  });
}

function loadPreparedBillingEntryMapFromSheet_(billingMonth) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const workbook = billingSs();
  if (!workbook) return {};
  const jsonSheet = workbook.getSheetByName('PreparedBillingJson');
  if (!monthKey || !jsonSheet) return {};

  const jsonLastRow = jsonSheet.getLastRow();
  if (jsonLastRow < 2) return {};

  const jsonValues = jsonSheet.getRange(2, 1, jsonLastRow - 1, 3).getValues();
  const entries = {};
  jsonValues.forEach(row => {
    if (String(row[0] || '').trim() !== monthKey) return;
    const jsonText = row[2];
    if (!jsonText) return;
    try {
      const parsedRow = JSON.parse(jsonText);
      const pid = billingNormalizePatientId_(parsedRow && parsedRow.patientId);
      if (!pid) return;
      entries[pid] = pickPreparedBillingEntrySummary_(parsedRow);
    } catch (err) {
      billingLogger_.log('[billing] loadPreparedBillingEntryMapFromSheet_ failed to parse billingRowJson for ' + monthKey + ': ' + err);
    }
  });
  return entries;
}

function pickPreparedBillingEntrySummary_(entry) {
  const row = entry || {};
  return {
    patientId: row.patientId || '',
    aggregateUntilMonth: row.aggregateUntilMonth || '',
    billingAmount: row.billingAmount,
    total: row.total,
    grandTotal: row.grandTotal,
    carryOverAmount: row.carryOverAmount,
    carryOverFromHistory: row.carryOverFromHistory,
    treatmentAmount: row.treatmentAmount,
    visitCount: row.visitCount,
    insuranceType: row.insuranceType,
    burdenRate: row.burdenRate,
    manualUnitPrice: row.manualUnitPrice,
    manualTransportAmount: row.manualTransportAmount,
    transportAmount: row.transportAmount,
    unitPrice: row.unitPrice,
    medicalAssistance: row.medicalAssistance,
    selfPayItems: row.selfPayItems,
    manualSelfPayAmount: row.manualSelfPayAmount
  };
}

function resolveTreatmentAmountForEntry_(entry) {
  if (!entry) return 0;
  const amountCalc = calculateBillingAmounts_({
    visitCount: entry.visitCount,
    insuranceType: entry.insuranceType,
    burdenRate: entry.burdenRate,
    manualUnitPrice: entry.manualUnitPrice != null ? entry.manualUnitPrice : entry.unitPrice,
    manualTransportAmount: Object.prototype.hasOwnProperty.call(entry, 'manualTransportAmount')
      ? entry.manualTransportAmount
      : entry.transportAmount,
    unitPrice: entry.unitPrice,
    medicalAssistance: entry.medicalAssistance,
    carryOverAmount: entry.carryOverAmount,
    selfPayItems: entry.selfPayItems,
    manualSelfPayAmount: entry.manualSelfPayAmount
  });

  return Number.isFinite(amountCalc && amountCalc.treatmentAmount) ? amountCalc.treatmentAmount : 0;
}

function resolveTreatmentAmountForMonthAndPatient_(billingMonth, patientId, fallbackEntry, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = billingNormalizePatientId_(patientId);
  if (!monthKey || !pid) return 0;

  if (fallbackEntry) {
    return resolveTreatmentAmountForEntry_(fallbackEntry);
  }

  const found = getPreparedBillingEntryForMonthCached_(monthKey, pid, cache);
  return resolveTreatmentAmountForEntry_(found);
}

function formatAggregateReceiptDescription_(months) {
  const normalized = normalizeReceiptMonthKeys_(months);
  if (!normalized.length) return '';
  const first = normalized[0];
  const last = normalized[normalized.length - 1];
  const firstLabel = `${first.slice(0, 4)}年${first.slice(4, 6)}月`;
  const lastLabel = `${last.slice(0, 4)}年${last.slice(4, 6)}月`;
  if (first === last) {
    return `${firstLabel}分 施術料金として`;
  }
  return `${firstLabel}〜${lastLabel}分 施術料金として`;
}

function formatAggregateReceiptMonthLabel_(monthKey) {
  const normalized = normalizeBillingMonthKeySafe_(monthKey);
  if (!normalized) return '';
  const year = normalized.slice(0, 4);
  const month = normalized.slice(4, 6);
  return `${year}年${month}月 施術料`;
}

function resolveAggregateReceiptForEntry_(patientId, previousMonthKey, prepared, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(previousMonthKey);
  const pid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(patientId)
    : String(patientId || '').trim();
  if (!monthKey || !pid) return null;

  const normalized = getPreparedBillingForMonthCached_(monthKey, cache);
  if (!normalized) return null;
  const previousEntry = getPreparedBillingEntryForMonthCached_(monthKey, pid, cache);

  const aggregateMonths = resolveAggregateMonthsFromUnpaid_(
    normalizeBillingMonthKeySafe_(prepared && prepared.billingMonth) || monthKey,
    pid,
    { useLegacyAggregate: false },
    prepared || normalized,
    cache
  );
  if (!aggregateMonths.length) return null;

  const breakdown = [];
  let total = 0;
  aggregateMonths.forEach(month => {
    const amount = resolveTreatmentAmountForMonthAndPatient_(
      month,
      pid,
      month === monthKey ? previousEntry : null,
      cache
    );
    total += amount;
    const label = formatAggregateReceiptMonthLabel_(month);
    breakdown.push({ month: label || month, amount });
  });

  return {
    months: aggregateMonths,
    breakdown,
    total,
    remark: formatAggregateReceiptDescription_(aggregateMonths)
  };
}

function buildReceiptMonthsFromBankUnpaid_(patientId, anchorMonthKey, prepared, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(anchorMonthKey);
  const pid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(patientId)
    : String(patientId || '').trim();

  if (!monthKey || !pid) return [];

  const workbook = billingSs();
  const months = [];
  let cursor = monthKey;
  let guard = 0;

  while (cursor && guard < 48) {
    const sheet = workbook.getSheetByName(formatBankWithdrawalSheetName_(cursor));
    if (!sheet) break;

    months.push(cursor);

    const isChecked = isPatientCheckedUnpaidInBankWithdrawalSheet_(pid, cursor, prepared, cache);
    if (!isChecked) break;
    cursor = resolvePreviousBillingMonthKey_(cursor);
    guard += 1;
  }

  return months.reverse();
}

// Internal helper: build map of unpaid patients for the month (not a boolean predicate).
function collectBankWithdrawalUnpaidPatients_(billingMonth, prepared) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) return {};

  const workbook = billingSs();
  const sheetName = formatBankWithdrawalSheetName_(monthKey);
  const sheet = workbook.getSheetByName(sheetName);
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const unpaidCol = resolveBillingColumn_(headers, [BANK_WITHDRAWAL_UNPAID_HEADER], BANK_WITHDRAWAL_UNPAID_HEADER, {});
  if (!unpaidCol) return {};
  const pidCol = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', {});
  const nameCol = resolveBillingColumn_(headers, BILLING_LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const kanaCol = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', {});
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const patients = resolvePreparedPatients_(prepared);
  const nameToPatientId = buildPatientNameToIdMap_(patients);
  const normalizePid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_
    : value => String(value || '').trim();

  const unpaidMap = {};
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    if (!row[unpaidCol - 1]) continue;

    const resolvedPid = pidCol
      ? normalizePid(row[pidCol - 1])
      : nameToPatientId[buildFullNameKey_(row[nameCol - 1], kanaCol ? row[kanaCol - 1] : '')];

    if (resolvedPid) unpaidMap[resolvedPid] = true;
  }

  return unpaidMap;
}

function collectBankWithdrawalStatusByPatient_(billingMonth) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) return {};

  const workbook = billingSs();
  const sheetName = formatBankWithdrawalSheetName_(monthKey);
  const sheet = workbook.getSheetByName(sheetName);
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const unpaidCol = resolveBillingColumn_(headers, [BANK_WITHDRAWAL_UNPAID_HEADER], BANK_WITHDRAWAL_UNPAID_HEADER, {});
  const aggregateCol = resolveBillingColumn_(headers, [BANK_WITHDRAWAL_AGGREGATE_HEADER], BANK_WITHDRAWAL_AGGREGATE_HEADER, {});
  const amountCol = resolveBillingColumn_(
    headers,
    ['金額', '請求金額', '引落額', '引落金額'],
    '金額',
    { fallbackLetter: BANK_WITHDRAWAL_AMOUNT_COLUMN_LETTER }
  );
  const pidCol = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', {});
  if (!pidCol) return {};

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const normalizePid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_
    : value => String(value || '').trim();
  const normalizeAmount = typeof normalizeMoneyNumber_ === 'function'
    ? normalizeMoneyNumber_
    : value => Number(value) || 0;
  const statusByPatient = {};

  values.forEach(row => {
    const pid = normalizePid(row[pidCol - 1]);
    if (!pid) return;

    const ae = unpaidCol ? normalizeBankFlagValue_(row[unpaidCol - 1]) : false;
    const af = aggregateCol ? normalizeBankFlagValue_(row[aggregateCol - 1]) : false;
    const amount = amountCol ? normalizeAmount(row[amountCol - 1]) : 0;
    const current = statusByPatient[pid] || { ae: false, af: false, amount: 0 };
    statusByPatient[pid] = {
      ae: current.ae || ae,
      af: current.af || af,
      amount: current.amount + (Number.isFinite(amount) ? amount : 0)
    };
  });

  return statusByPatient;
}

function getBankWithdrawalStatusByPatient_(billingMonth, patientId, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(patientId)
    : String(patientId || '').trim();

  if (!monthKey || !pid) return { ae: false, af: false, amount: 0 };

  const store = cache && cache.bankWithdrawalStatusByMonth
    ? cache.bankWithdrawalStatusByMonth
    : null;
  const statusByPatient = store
    ? (Object.prototype.hasOwnProperty.call(store, monthKey)
      ? store[monthKey]
      : (store[monthKey] = collectBankWithdrawalStatusByPatient_(monthKey) || {}))
    : collectBankWithdrawalStatusByPatient_(monthKey);

  const entry = statusByPatient && statusByPatient[pid];
  if (!entry) return { ae: false, af: false, amount: 0 };

  return {
    ae: !!entry.ae,
    af: !!entry.af,
    amount: Number(entry.amount) || 0
  };
}

function getBankWithdrawalStatusByPatient(billingMonth, patientId) {
  return getBankWithdrawalStatusByPatient_(billingMonth, patientId);
}

function resolveInvoiceModeFromBankFlags_(billingMonth, patientId, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(patientId)
    : String(patientId || '').trim();

  if (!monthKey || !pid) {
    return { mode: 'standard', months: [] };
  }

  const flags = getBankWithdrawalStatusByPatient_(monthKey, pid, cache);
  const ae = !!(flags && flags.ae);
  const af = !!(flags && flags.af);

  if (ae || af) {
    return { mode: 'skip', months: [] };
  }

  const previousMonthKey = resolvePreviousBillingMonthKey_(monthKey);
  if (!previousMonthKey) {
    return { mode: 'standard', months: [] };
  }

  const previousFlags = getBankWithdrawalStatusByPatient_(previousMonthKey, pid, cache);
  if (!(previousFlags && previousFlags.af)) {
    return { mode: 'standard', months: [] };
  }

  const unpaidMonths = collectAggregateBankFlagMonthsForPatient_(previousMonthKey, pid, null, cache);
  const months = normalizePastBillingMonths_(unpaidMonths.concat(previousMonthKey), monthKey);
  return { mode: 'aggregate', months };
}

function getBankWithdrawalUnpaidMapCached_(billingMonth, prepared, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) return {};
  const store = cache && cache.bankWithdrawalUnpaidByMonth ? cache.bankWithdrawalUnpaidByMonth : null;
  if (store && Object.prototype.hasOwnProperty.call(store, monthKey)) {
    return store[monthKey] || {};
  }
  const unpaidMap = collectBankWithdrawalUnpaidPatients_(monthKey, prepared);
  if (store) {
    store[monthKey] = unpaidMap || {};
  }
  return unpaidMap || {};
}

function isPatientCheckedUnpaidInBankWithdrawalSheet_(patientId, billingMonth, prepared, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(patientId)
    : String(patientId || '').trim();

  if (!monthKey || !pid) return false;
  const unpaidMap = getBankWithdrawalUnpaidMapCached_(monthKey, prepared, cache);
  return !!(unpaidMap && unpaidMap[pid]);
}

function attachPreviousReceiptAmounts_(prepared, cache) {
  const monthKey = prepared && prepared.billingMonth;
  if (!monthKey || !Array.isArray(prepared && prepared.billingJson)) return prepared;

  const previousMonthKey = resolvePreviousBillingMonthKey_(monthKey);
  if (!previousMonthKey) return prepared;

  const normalizePid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_
    : value => String(value || '').trim();
  const monthCache = cache || {
    preparedByMonth: {},
    bankWithdrawalUnpaidByMonth: {},
    bankWithdrawalAmountsByMonth: {}
  };
  const previousPrepared = getPreparedBillingForMonthCached_(previousMonthKey, monthCache);
  if (monthCache.preparedByMonth) {
    monthCache.preparedByMonth[monthKey] = monthCache.preparedByMonth[monthKey] || prepared;
    if (previousMonthKey) {
      monthCache.preparedByMonth[previousMonthKey] = monthCache.preparedByMonth[previousMonthKey] || previousPrepared;
    }
  }

  const enrichedJson = prepared.billingJson.map(entry => {
    const pid = normalizePid(entry && entry.patientId);
    const hasPreviousPreparedEntry = !!(previousPrepared && getPreparedBillingEntryForPatient_(previousPrepared, pid));
    const receiptTargetMonths = resolveReceiptTargetMonthsFromBankFlags_(pid, monthKey, prepared, monthCache);
    const hasPreviousReceiptSheet = hasPreviousPreparedEntry;
    const currentFlags = prepared && prepared.bankFlagsByPatient && prepared.bankFlagsByPatient[pid];
    const previousFlags = previousPrepared && previousPrepared.bankFlagsByPatient && previousPrepared.bankFlagsByPatient[pid];

    if (receiptTargetMonths.length > 1) {
      const receiptBreakdown = buildReceiptMonthBreakdownForEntry_(pid, receiptTargetMonths, previousPrepared || prepared, monthCache);
      const aggregateRemark = formatAggregateReceiptDescription_(receiptTargetMonths);
      const previousReceipt = entry && entry.previousReceipt ? entry.previousReceipt : {
        addressee: '株式会社べるつりー',
        note: aggregateRemark
      };
      logReceiptDebug_(pid, {
        step: 'attachPreviousReceiptAmounts_',
        billingMonth: monthKey,
        patientId: pid,
        currentFlags,
        previousFlags,
        receiptTargetMonths,
        receiptMonths: receiptTargetMonths
      });
      return Object.assign({}, entry, {
        hasPreviousReceiptSheet,
        receiptMonths: receiptTargetMonths,
        receiptRemark: aggregateRemark,
        receiptMonthBreakdown: receiptBreakdown,
        previousReceipt
      });
    }

    if (!receiptTargetMonths.length) {
      logReceiptDebug_(pid, {
        step: 'attachPreviousReceiptAmounts_',
        billingMonth: monthKey,
        patientId: pid,
        currentFlags,
        previousFlags,
        receiptTargetMonths,
        receiptMonths: []
      });
      return Object.assign({}, entry, {
        hasPreviousReceiptSheet,
        receiptMonths: [],
        receiptRemark: '',
        receiptMonthBreakdown: hasPreviousReceiptSheet ? (entry && entry.receiptMonthBreakdown) : []
      });
    }

    if (!hasPreviousReceiptSheet && receiptTargetMonths[0] === previousMonthKey) {
      return Object.assign({}, entry, {
        hasPreviousReceiptSheet: false,
        receiptRemark: '',
        receiptMonthBreakdown: []
      });
    }

    const resolveLegacyPreviousReceiptAmount = () => {
      if (!previousPrepared || !Array.isArray(previousPrepared.billingJson)) return 0;
      const legacyEntry = previousPrepared.billingJson.find(item => normalizePid(item && item.patientId) === pid);
      if (!legacyEntry) return 0;
      if (legacyEntry.grandTotal != null && legacyEntry.grandTotal !== '') {
        return normalizeMoneyNumber_(legacyEntry.grandTotal);
      }
      return normalizeMoneyNumber_(legacyEntry.billingAmount);
    };

    const receiptBreakdown = receiptTargetMonths.length === 1
      && receiptTargetMonths[0] === previousMonthKey
      && !hasPreviousReceiptSheet
      ? [{ month: previousMonthKey, amount: resolveLegacyPreviousReceiptAmount() }]
      : buildReceiptMonthBreakdownForEntry_(pid, receiptTargetMonths, previousPrepared || prepared, monthCache);

    logReceiptDebug_(pid, {
      step: 'attachPreviousReceiptAmounts_',
      billingMonth: monthKey,
      patientId: pid,
      currentFlags,
      previousFlags,
      receiptTargetMonths,
      receiptMonths: receiptTargetMonths
    });
    return Object.assign({}, entry, {
      hasPreviousReceiptSheet,
      receiptMonths: receiptTargetMonths,
      receiptRemark: entry && entry.receiptRemark,
      receiptMonthBreakdown: receiptBreakdown
    });
  });

  return Object.assign({}, prepared, {
    billingJson: enrichedJson,
    hasPreviousReceiptSheet: !!previousPrepared
  });
}

function collectPreviousReceiptAmountsFromBankSheet_(billingMonth, prepared) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) return { hasSheet: false, amounts: {} };

  const workbook = billingSs();
  const sheetName = formatBankWithdrawalSheetName_(monthKey);
  const sheet = workbook.getSheetByName(sheetName);
  if (!sheet) return { hasSheet: false, amounts: {} };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { hasSheet: true, amounts: {} };

  const amountColIndex = columnLetterToNumber_(BANK_WITHDRAWAL_AMOUNT_COLUMN_LETTER);
  const headerCount = Math.max(sheet.getLastColumn(), amountColIndex || 0);
  const headers = sheet.getRange(1, 1, 1, headerCount).getDisplayValues()[0];
  const amountCol = columnLetterToNumber_(BANK_WITHDRAWAL_AMOUNT_COLUMN_LETTER);
  const pidCol = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', {});
  const nameCol = resolveBillingColumn_(headers, BILLING_LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const kanaCol = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', {});
  const values = sheet.getRange(2, 1, lastRow - 1, headerCount).getValues();

  const normalizePid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_
    : value => String(value || '').trim();
  const patients = resolvePreparedPatients_(prepared);
  const nameToPatientId = buildPatientNameToIdMap_(patients);
  const amounts = {};

  values.forEach(row => {
    const resolvedPid = pidCol
      ? normalizePid(row[pidCol - 1])
      : nameToPatientId[buildFullNameKey_(row[nameCol - 1], kanaCol ? row[kanaCol - 1] : '')];
    if (!resolvedPid) return;

    const amount = amountCol ? Number(row[amountCol - 1]) || 0 : 0;
    if (!amount) return;

    amounts[resolvedPid] = (amounts[resolvedPid] || 0) + amount;
  });

  return { hasSheet: true, amounts };
}

function collectBankWithdrawalAmountsByPatientCached_(billingMonth, prepared, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) return {};
  const store = cache && cache.bankWithdrawalAmountsByMonth ? cache.bankWithdrawalAmountsByMonth : (cache || {});
  if (!Object.prototype.hasOwnProperty.call(store, monthKey)) {
    store[monthKey] = collectBankWithdrawalAmountsByPatient_(monthKey, prepared) || {};
  }
  return store[monthKey];
}

function buildReceiptMonthBreakdownForEntry_(patientId, months, prepared, cache) {
  const normalizePid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_
    : value => String(value || '').trim();
  const pid = normalizePid(patientId);
  if (!pid || !Array.isArray(months) || !months.length) return [];

  const seen = new Set();
  const breakdown = [];
  const store = cache || {};
  const preparedMonthKey = normalizeBillingMonthKeySafe_(prepared && prepared.billingMonth);

  months.forEach(month => {
    const monthKey = normalizeBillingMonthKeySafe_(month);
    if (!monthKey || seen.has(monthKey)) return;
    seen.add(monthKey);

    if (preparedMonthKey && preparedMonthKey === monthKey && Array.isArray(prepared && prepared.billingJson)) {
      const match = prepared.billingJson.find(item => normalizePid(item && item.patientId) === pid);
      if (match) {
        const amount = match.grandTotal != null && match.grandTotal !== ''
          ? normalizeMoneyNumber_(match.grandTotal)
          : normalizeMoneyNumber_(match.billingAmount);
        if (Number.isFinite(amount)) {
          breakdown.push({ month: monthKey, amount });
          return;
        }
      }
    }

    const amountByPatient = collectBankWithdrawalAmountsByPatientCached_(monthKey, prepared, store);
    if (amountByPatient && Object.prototype.hasOwnProperty.call(amountByPatient, pid)) {
      const amount = amountByPatient[pid];
      if (Number.isFinite(amount)) {
        breakdown.push({ month: monthKey, amount });
        return;
      }
    }

    if (preparedMonthKey && Number(monthKey) < Number(preparedMonthKey)) {
      const previousPrepared = getPreparedBillingForMonthCached_(monthKey, store);
      if (previousPrepared && Array.isArray(previousPrepared.billingJson)) {
        const match = previousPrepared.billingJson.find(item => normalizePid(item && item.patientId) === pid);
        if (match) {
          const amount = match.grandTotal != null && match.grandTotal !== ''
            ? normalizeMoneyNumber_(match.grandTotal)
            : normalizeMoneyNumber_(match.billingAmount);
          if (Number.isFinite(amount)) {
            breakdown.push({ month: monthKey, amount });
          }
        }
      }
    }
  });

  return breakdown;
}

function collectBankWithdrawalAmountsByPatient_(billingMonth, prepared) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) return {};

  const month = normalizeBillingMonthInput(monthKey);
  const workbook = billingSs();
  const sheetName = formatBankWithdrawalSheetName_(month);
  const sheet = workbook.getSheetByName(sheetName);
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const unpaidCol = resolveBillingColumn_(headers, [BANK_WITHDRAWAL_UNPAID_HEADER], BANK_WITHDRAWAL_UNPAID_HEADER, {});
  const amountCol = resolveBillingColumn_(
    headers,
    ['金額', '請求金額', '引落額', '引落金額'],
    '金額',
    { required: true, fallbackLetter: BANK_WITHDRAWAL_AMOUNT_COLUMN_LETTER }
  );
  const pidCol = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', {});
  const nameCol = resolveBillingColumn_(headers, BILLING_LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const kanaCol = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', {});
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const patients = resolvePreparedPatients_(prepared);
  const nameToPatientId = buildPatientNameToIdMap_(patients);
  const normalizePid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_
    : value => String(value || '').trim();
  const amounts = {};

  values.forEach(row => {
    if (unpaidCol && row[unpaidCol - 1]) return;
    const amount = amountCol ? Number(row[amountCol - 1]) || 0 : 0;
    if (!amount) return;

    const resolvedPid = pidCol
      ? normalizePid(row[pidCol - 1])
      : nameToPatientId[buildFullNameKey_(row[nameCol - 1], kanaCol ? row[kanaCol - 1] : '')];

    if (!resolvedPid) return;
    amounts[resolvedPid] = (amounts[resolvedPid] || 0) + amount;
  });
  return amounts;
}

function resolveBillingAmountForEntry_(entry) {
  if (!entry) return 0;
  const carryOverTotal = normalizeMoneyNumber_(entry.carryOverAmount)
    + normalizeMoneyNumber_(entry.carryOverFromHistory);

  if (entry.grandTotal != null && entry.grandTotal !== '') {
    return normalizeMoneyNumber_(entry.grandTotal);
  }

  if (entry.total != null && entry.total !== '') {
    return normalizeMoneyNumber_(entry.total) + carryOverTotal;
  }

  const billingAmount = normalizeMoneyNumber_(entry.billingAmount);
  const transportAmount = normalizeMoneyNumber_(entry.transportAmount);

  if (entry.billingAmount != null || entry.transportAmount != null || carryOverTotal) {
    return billingAmount + transportAmount + carryOverTotal;
  }

  return 0;
}

function resolveBillingAmountForMonthAndPatient_(billingMonth, patientId, fallbackEntry, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = billingNormalizePatientId_(patientId);
  if (!monthKey || !pid) return 0;

  if (fallbackEntry) {
    return resolveBillingAmountForEntry_(fallbackEntry);
  }

  const normalized = getPreparedBillingForMonthCached_(monthKey, cache);
  if (!normalized) return 0;

  const totalsEntry = normalized.totalsByPatient && normalized.totalsByPatient[pid];
  if (totalsEntry) {
    return resolveBillingAmountForEntry_(totalsEntry);
  }

  const found = getPreparedBillingEntryForMonthCached_(monthKey, pid, cache);
  return resolveBillingAmountForEntry_(found);
}

function collectAggregateBankFlagMonthsForPatient_(billingMonth, patientId, aggregateUntilMonth, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = billingNormalizePatientId_(patientId);
  if (!monthKey || !pid) return [];

  const months = [];
  let cursor = resolvePreviousBillingMonthKey_(monthKey);
  let guard = 0;

  while (cursor && guard < 48) {
    const cursorKey = normalizeBillingMonthKeySafe_(cursor);
    if (!cursorKey) break;

    const normalized = getPreparedBillingForMonthCached_(cursorKey, cache);
    // 安全側仕様: 中間月で prepared/bankFlags が欠損している場合は履歴走査を打ち切る
    // （不確実な状態での自動合算を防ぐ）
    if (!normalized || !normalized.bankFlagsByPatient) break;

    const preparedEntry = getPreparedBillingEntryForPatient_(normalized, pid);
    if (!preparedEntry) break;

    const flags = normalized.bankFlagsByPatient[pid];
    if (flags && flags.ae) {
      months.unshift(normalized.billingMonth || cursorKey);
      cursor = resolvePreviousBillingMonthKey_(normalized.billingMonth || cursorKey);
      guard += 1;
      continue;
    }
    break;
  }

  return months;
}

function resolveReceiptTargetMonthsFromBankFlags_(patientId, currentMonth, prepared, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(currentMonth);
  const pid = billingNormalizePatientId_(patientId);
  if (!monthKey || !pid) return [];

  const previousMonthKey = resolvePreviousBillingMonthKey_(monthKey);
  if (!previousMonthKey) return [];

  const previousPrepared = getPreparedBillingForMonthCached_(previousMonthKey, cache);
  const previousFlags = previousPrepared && previousPrepared.bankFlagsByPatient && previousPrepared.bankFlagsByPatient[pid];
  const currentFlags = prepared && prepared.bankFlagsByPatient && prepared.bankFlagsByPatient[pid];

  if (currentFlags && (currentFlags.ae || currentFlags.af)) {
    logReceiptDebug_(pid, {
      step: 'resolveReceiptTargetMonthsFromBankFlags_',
      billingMonth: monthKey,
      patientId: pid,
      currentFlags,
      previousFlags,
      receiptTargetMonths: []
    });
    return [];
  }

  if (previousFlags && previousFlags.af) {
    const unpaidMonths = collectAggregateBankFlagMonthsForPatient_(previousMonthKey, pid, null, cache);
    const receiptTargetMonths = normalizePastBillingMonths_(unpaidMonths.concat(previousMonthKey), monthKey);
    logReceiptDebug_(pid, {
      step: 'resolveReceiptTargetMonthsFromBankFlags_',
      billingMonth: monthKey,
      patientId: pid,
      currentFlags,
      previousFlags,
      receiptTargetMonths
    });
    return receiptTargetMonths;
  }

  if (previousFlags && previousFlags.ae) {
    logReceiptDebug_(pid, {
      step: 'resolveReceiptTargetMonthsFromBankFlags_',
      billingMonth: monthKey,
      patientId: pid,
      currentFlags,
      previousFlags,
      receiptTargetMonths: []
    });
    return [];
  }

  const receiptTargetMonths = normalizePastBillingMonths_([previousMonthKey], monthKey);
  logReceiptDebug_(pid, {
    step: 'resolveReceiptTargetMonthsFromBankFlags_',
    billingMonth: monthKey,
    patientId: pid,
    currentFlags,
    previousFlags,
    receiptTargetMonths
  });
  return receiptTargetMonths;
}

function formatAggregateBillingRemark_(months) {
  const normalized = (Array.isArray(months) ? months : [])
    .map(value => normalizeBillingMonthKeySafe_(value))
    .filter(Boolean);
  const seen = new Set();
  const unique = [];
  normalized.forEach(ym => {
    if (!seen.has(ym)) {
      seen.add(ym);
      unique.push(ym);
    }
  });

  if (!unique.length) return '';

  const labels = unique.map((ym, idx) => {
    const year = ym.slice(0, 4);
    const month = ym.slice(4, 6);
    return idx === 0 ? `${year}年${month}月分` : `${month}月分`;
  });

  return labels.join('・') + ' 合算請求';
}

function hasAggregatedThisUnpaidCycle_(entry, flags, unpaidSet, monthKey) {
  if (!entry || !Array.isArray(entry.aggregateTargetMonths) || !entry.aggregateTargetMonths.length) return false;
  if (!flags || flags.ae) return false;
  const normalizedMonthKey = normalizeBillingMonthKeySafe_(monthKey);
  const normalizedTargets = entry.aggregateTargetMonths.map(normalizeBillingMonthKeySafe_).filter(Boolean);
  return normalizedMonthKey && normalizedTargets.indexOf(normalizedMonthKey) >= 0
    && normalizedTargets.some(target => unpaidSet.has(target));
}

function resolveAggregateMonthsFromUnpaid_(billingMonth, patientId, options, prepared, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = billingNormalizePatientId_(patientId);
  if (!monthKey || !pid) return [];

  const previousMonthKey = resolvePreviousBillingMonthKey_(monthKey);
  if (!previousMonthKey) return [];

  const previousPrepared = getPreparedBillingForMonthCached_(previousMonthKey, cache);
  const previousFlags = previousPrepared && previousPrepared.bankFlagsByPatient && previousPrepared.bankFlagsByPatient[pid];
  const currentFlags = prepared && prepared.bankFlagsByPatient && prepared.bankFlagsByPatient[pid];

  if (currentFlags && (currentFlags.ae || currentFlags.af)) return [];
  if (!(previousFlags && previousFlags.af)) return [];

  const unpaidMonths = collectAggregateBankFlagMonthsForPatient_(previousMonthKey, pid, null, cache);
  return normalizePastBillingMonths_(unpaidMonths.concat(previousMonthKey), monthKey);
}

function sanitizeAggregateFieldsForBankFlags_(entry, bankFlags) {
  const flags = bankFlags || {};
  if (flags.af === true) return entry;
  const sanitized = Object.assign({}, entry || {});
  delete sanitized.aggregateStatus;
  delete sanitized.aggregateRemark;
  delete sanitized.receiptMonths;
  delete sanitized.aggregateTargetMonths;
  return sanitized;
}

function applyAggregateInvoiceRulesFromBankFlags_(prepared, cache) {
  const normalized = normalizePreparedBilling_(prepared);
  const monthKey = normalizeBillingMonthKeySafe_(normalized && normalized.billingMonth);
  if (!normalized || !monthKey || !Array.isArray(normalized.billingJson)) return normalized || prepared;

  const bankFlagsByPatient = normalized.bankFlagsByPatient || {};
  const monthCache = cache || {
    preparedByMonth: {},
    bankWithdrawalUnpaidByMonth: {},
    bankWithdrawalAmountsByMonth: {}
  };
  const previousMonthKey = resolvePreviousBillingMonthKey_(monthKey);
  const previousPrepared = previousMonthKey ? getPreparedBillingForMonthCached_(previousMonthKey, monthCache) : null;

  const transformed = normalized.billingJson.map(entry => {
    const pid = billingNormalizePatientId_(entry && entry.patientId);
    if (!pid) return entry;
    const flags = bankFlagsByPatient && Object.prototype.hasOwnProperty.call(bankFlagsByPatient, pid)
      ? bankFlagsByPatient[pid]
      : null;
    const previousFlags = previousPrepared && previousPrepared.bankFlagsByPatient && previousPrepared.bankFlagsByPatient[pid];
    if (flags && (flags.ae || flags.af)) {
      return sanitizeAggregateFieldsForBankFlags_(entry, flags);
    }
    if (!(previousFlags && previousFlags.af)) {
      return sanitizeAggregateFieldsForBankFlags_(entry, flags);
    }

    const targetMonths = resolveAggregateMonthsFromUnpaid_(monthKey, pid, {}, normalized, monthCache);
    if (targetMonths.length <= 1) {
      logReceiptDebug_(pid, {
        step: 'applyAggregateInvoiceRulesFromBankFlags_',
        billingMonth: monthKey,
        patientId: pid,
        currentFlags: flags,
        previousFlags,
        receiptTargetMonths: targetMonths,
        receiptMonths: []
      });
      return Object.assign({}, entry, {
        receiptMonths: []
      });
    }
    const aggregateTotal = targetMonths.reduce(
      (sum, ym) => sum + resolveBillingAmountForMonthAndPatient_(ym, pid, null, monthCache),
      0
    );
    const aggregateRemark = formatAggregateBillingRemark_(targetMonths);

    logReceiptDebug_(pid, {
      step: 'applyAggregateInvoiceRulesFromBankFlags_',
      billingMonth: monthKey,
      patientId: pid,
      currentFlags: flags,
      previousFlags,
      receiptTargetMonths: targetMonths,
      receiptMonths: targetMonths
    });
    return Object.assign({}, entry, {
      aggregateRemark,
      receiptMonths: targetMonths,
      billingAmount: aggregateTotal,
      transportAmount: 0,
      total: aggregateTotal,
      grandTotal: aggregateTotal,
      aggregateTargetMonths: targetMonths,
      skipReceipt: false
    });
  });

  return Object.assign({}, normalized, { billingJson: transformed });
}

function savePreparedBillingJsonRows_(billingMonth, billingJson) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey || !Array.isArray(billingJson)) return { billingMonth: monthKey || '', inserted: 0 };

  const sheet = ensurePreparedBillingJsonSheet_();
  if (!sheet) return { billingMonth: monthKey, inserted: 0 };
  const lastRow = sheet.getLastRow();
  const existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];

  for (let i = existingRows.length - 1; i >= 0; i--) {
    if (String(existingRows[i][0] || '').trim() === monthKey) {
      sheet.deleteRow(i + 2);
    }
  }

  if (!billingJson.length) return { billingMonth: monthKey, inserted: 0 };

  const normalizePid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_
    : value => String(value || '').trim();

  const rows = billingJson.map(entry => {
    const rowPayload = entry || {};
    const pid = normalizePid(rowPayload.patientId);
    return [monthKey, pid || '', JSON.stringify(rowPayload)];
  });

  sheet.insertRows(2, rows.length);
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  return { billingMonth: monthKey, inserted: rows.length };
}

function savePreparedBillingToSheet_(billingMonth, preparedPayload) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth || (preparedPayload && preparedPayload.billingMonth));
  const normalized = normalizePreparedBilling_(preparedPayload);
  if (!monthKey || !normalized) {
    billingLogger_.log('[billing] savePreparedBillingToSheet_ skipped due to invalid payload');
    return null;
  }

  const payload = Object.assign({}, normalized, { billingMonth: monthKey });
  const metaPayload = Object.assign({}, payload);
  delete metaPayload.billingJson;
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
  const metaResult = savePreparedBillingMeta_(monthKey, {
    preparedAt: preparedAtValue,
    preparedBy,
    payloadVersion,
    metaPayload,
    note: ''
  });
  const jsonResult = savePreparedBillingJsonRows_(monthKey, Array.isArray(payload.billingJson) ? payload.billingJson : []);
  return { billingMonth: monthKey, meta: metaResult, json: jsonResult };
}

function loadPreparedBillingFromSheet_(billingMonth) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const workbook = billingSs();
  if (!workbook) return null;
  const metaSheet = workbook.getSheetByName('PreparedBillingMeta');
  if (!monthKey || !metaSheet) return null;

  const metaLastRow = metaSheet.getLastRow();
  if (metaLastRow < 2) return null;

  const metaValues = metaSheet.getRange(2, 1, metaLastRow - 1, 5).getValues();
  const metaRow = metaValues.find(row => String(row[0] || '').trim() === monthKey);
  if (!metaRow) return null;

  const preparedAtCell = metaRow[1];
  const preparedByCell = metaRow[2];
  const payloadVersion = metaRow[3];
  let parsed = {};
  const metaJsonSheet = workbook.getSheetByName('PreparedBillingMetaJson');
  if (metaJsonSheet) {
    const metaLast = metaJsonSheet.getLastRow();
    if (metaLast >= 2) {
      const rows = metaJsonSheet.getRange(2, 1, metaLast - 1, 3).getValues();
      const chunks = rows
        .filter(row => String(row[0] || '').trim() === monthKey)
        .sort((a, b) => Number(a[1]) - Number(b[1]))
        .map(row => row[2] || '')
        .filter(Boolean);
      const payloadJson = chunks.join('');
      if (payloadJson) {
        try {
          parsed = JSON.parse(payloadJson) || {};
        } catch (err) {
          billingLogger_.log('[billing] loadPreparedBillingFromSheet_ failed to parse meta payload for ' + monthKey + ': ' + err);
        }
      }
    }
  }

  if (!parsed || typeof parsed !== 'object') {
    parsed = {};
  }

  const jsonSheet = workbook.getSheetByName('PreparedBillingJson');
  const billingJsonRows = [];
  if (jsonSheet) {
    const jsonLastRow = jsonSheet.getLastRow();
    if (jsonLastRow >= 2) {
      const jsonValues = jsonSheet.getRange(2, 1, jsonLastRow - 1, 3).getValues();
      jsonValues.forEach(row => {
        if (String(row[0] || '').trim() !== monthKey) return;
        const jsonText = row[2];
        if (!jsonText) return;
        try {
          const parsedRow = JSON.parse(jsonText);
          billingJsonRows.push(parsedRow);
        } catch (err) {
          billingLogger_.log('[billing] loadPreparedBillingFromSheet_ failed to parse billingRowJson for ' + monthKey + ': ' + err);
        }
      });
    }
  }

  const merged = Object.assign({}, parsed, {
    billingMonth: monthKey,
    billingJson: billingJsonRows,
    preparedAt: parsed.preparedAt || (preparedAtCell instanceof Date ? preparedAtCell.toISOString() : null),
    preparedBy: parsed.preparedBy || preparedByCell || '',
    schemaVersion: parsed.schemaVersion || payloadVersion || PREPARED_BILLING_SCHEMA_VERSION
  });

  const normalized = normalizePreparedBilling_(merged);
  return normalized;
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

  const billingJson = payload.billingJson;
  if (!Array.isArray(billingJson)) return { ok: false, reason: 'billingJson missing' };
  if (billingJson.length === 0) return { ok: false, reason: 'billingJson empty' };

  const requiredArrays = [
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

// Deprecated: Prepared billing should be read from the PreparedBilling sheet for durability.
// This cache accessor is retained for backward compatibility with legacy flows.
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
  const cached = loadBillingCachePayload_(cache, key);
  if (!cached) {
    try {
      if (isBillingDebugEnabled_()) {
        Logger.log('[billing] loadPreparedBilling_: cache miss for ' + key);
      }
    } catch (err) {
      // ignore logging errors in non-GAS environments
    }
    return wrapResult(null, { ok: false, reason: 'cache miss' });
  }
  try {
    if (isBillingDebugEnabled_()) {
      Logger.log('[billing] loadPreparedBilling_ raw cache for ' + key + ': ' + cached);
    }
  } catch (err) {
    // ignore logging errors in non-GAS environments
  }
  try {
    const parsed = JSON.parse(cached);
    const rawLength = Array.isArray(parsed && parsed.billingJson) ? parsed.billingJson.length : 0;
    const validation = validatePreparedBillingPayload_(parsed, expectedMonthKey);
    const normalized = normalizePreparedBilling_(parsed);
    const normalizedLength = normalized && Array.isArray(normalized.billingJson) ? normalized.billingJson.length : 0;
    if (isBillingDebugEnabled_()) {
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
    }
    if (!validation.ok) {
      if (validation.reason === 'billingMonth mismatch' && normalized && expectedMonthKey && Array.isArray(normalized.billingJson)) {
        const corrected = Object.assign({}, normalized, { billingMonth: expectedMonthKey });
        if (isBillingDebugEnabled_()) {
          billingLogger_.log('[billing] loadPreparedBilling_ auto-correcting billingMonth mismatch for cache key=' + key);
        }
        savePreparedBilling_(corrected);
        return wrapResult(corrected, validation);
      }
      if (allowInvalid && normalized && Array.isArray(normalized.billingJson)) {
        try {
          if (isBillingDebugEnabled_()) {
            Logger.log('[billing] loadPreparedBilling_: allowing invalid cache for ' + key + ' reason=' + validation.reason);
          }
        } catch (err) {
          // ignore logging errors in non-GAS environments
        }
        return wrapResult(Object.assign({}, normalized, { billingMonth: validation.billingMonth || expectedMonthKey }), validation);
      }
      try {
        if (isBillingDebugEnabled_()) {
          Logger.log('[billing] loadPreparedBilling_: invalid cache for ' + key + ' reason=' + validation.reason);
        }
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

function loadPreparedBillingWithSheetFallback_(billingMonthKey, options) {
  const opts = options || {};
  const withValidation = opts.withValidation === true;
  const allowInvalid = opts.allowInvalid === true;
  const restoreCache = opts.restoreCache !== false;
  const wrapResult = (payload, validation) => withValidation ? { prepared: payload, validation: validation || null } : payload;

  const cacheResult = loadPreparedBilling_(billingMonthKey, Object.assign({}, opts, { withValidation: true }));
  const cachePrepared = cacheResult && cacheResult.prepared !== undefined ? cacheResult.prepared : cacheResult;
  const cacheValidation = cacheResult && cacheResult.validation
    ? cacheResult.validation
    : (cacheResult && cacheResult.ok !== undefined ? cacheResult : null);
  const cacheOk = cacheValidation ? cacheValidation.ok : !!cachePrepared;

  if (cachePrepared && (cacheOk || allowInvalid)) {
    return wrapResult(cachePrepared, cacheValidation || null);
  }

  const month = normalizeBillingMonthInput ? normalizeBillingMonthInput(billingMonthKey) : null;
  const monthKey = month && month.key ? month.key : normalizeBillingMonthKeySafe_(billingMonthKey);
  const fromSheet = monthKey ? loadPreparedBillingFromSheet_(monthKey) : null;
  const sheetValidation = validatePreparedBillingPayload_(fromSheet, monthKey);
  const normalizedPrepared = normalizePreparedBilling_(fromSheet);

  if (normalizedPrepared && (sheetValidation.ok || allowInvalid)) {
    const preparedWithMonth = Object.assign({}, normalizedPrepared, { billingMonth: sheetValidation.billingMonth || monthKey });
    if (sheetValidation.ok && restoreCache && typeof savePreparedBilling_ === 'function') {
      try {
        savePreparedBilling_(preparedWithMonth);
        if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
          billingLogger_.log('[billing] loadPreparedBillingWithSheetFallback_ restored cache for ' + preparedWithMonth.billingMonth);
        }
      } catch (err) {
        // ignore caching/logging errors in non-GAS environments
      }
    }
    return wrapResult(preparedWithMonth, sheetValidation);
  }

  return wrapResult(null, sheetValidation);
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
    const serialized = JSON.stringify(payloadToCache);
    clearBillingCache_(key);
    if (serialized.length > BILLING_CACHE_MAX_ENTRY_LENGTH) {
      const chunkCount = Math.ceil(serialized.length / BILLING_CACHE_CHUNK_SIZE);
      for (let idx = 0; idx < chunkCount; idx++) {
        const chunkKey = buildBillingCacheChunkKey_(key, idx + 1);
        const chunkValue = serialized.slice(idx * BILLING_CACHE_CHUNK_SIZE, (idx + 1) * BILLING_CACHE_CHUNK_SIZE);
        cache.put(chunkKey, chunkValue, BILLING_CACHE_TTL_SECONDS);
      }
      cache.put(key, BILLING_CACHE_CHUNK_MARKER + chunkCount, BILLING_CACHE_TTL_SECONDS);
    } else {
      cache.put(key, serialized, BILLING_CACHE_TTL_SECONDS);
    }
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

function parseBillingCacheChunkCount_(value) {
  if (typeof value !== 'string') return 0;
  if (value.indexOf(BILLING_CACHE_CHUNK_MARKER) !== 0) return 0;
  const parsed = parseInt(value.slice(BILLING_CACHE_CHUNK_MARKER.length), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

function loadBillingCachePayload_(cache, key) {
  if (!cache || typeof cache.get !== 'function') return null;
  const cached = cache.get(key);
  const chunkCount = parseBillingCacheChunkCount_(cached);
  if (!chunkCount) return cached;

  const chunks = [];
  for (let idx = 1; idx <= chunkCount; idx++) {
    const chunkValue = cache.get(buildBillingCacheChunkKey_(key, idx));
    chunks.push(chunkValue || '');
  }

  const merged = chunks.join('');
  if (!merged) return null;
  try {
    billingLogger_.log('[billing] loadBillingCachePayload_ merged chunked cache for ' + key + ' chunks=' + chunkCount);
  } catch (err) {
    // ignore logging errors
  }
  return merged;
}

function getPreparedBillingEntryForPatient_(prepared, patientId) {
  if (!prepared || !Array.isArray(prepared.billingJson)) return null;
  const pid = billingNormalizePatientId_(patientId);
  if (!pid) return null;
  return prepared.billingJson.find(entry => billingNormalizePatientId_(entry && entry.patientId) === pid) || null;
}

function getActiveUserEmail_() {
  try {
    const active = Session.getActiveUser && Session.getActiveUser();
    if (active && typeof active.getEmail === 'function') {
      return active.getEmail() || '';
    }
  } catch (err) {
    // ignore
  }
  return '';
}

function normalizeEmailKeySafe_(email) {
  return String(email || '')
    .trim()
    .toLowerCase();
}

function getBillingAdminEmails_() {
  try {
    const scriptProps = typeof PropertiesService !== 'undefined'
      ? PropertiesService.getScriptProperties()
      : null;
    const configured = scriptProps && scriptProps.getProperty('BILLING_ADMIN_EMAILS');
    if (configured && typeof configured === 'string') {
      return configured
        .split(/[\s,]+/)
        .map(normalizeEmailKeySafe_)
        .filter(Boolean);
    }
  } catch (err) {
    // ignore property read errors
  }

  try {
    if (typeof APP !== 'undefined' && APP.BILLING_ADMIN_EMAILS) {
      const appValue = APP.BILLING_ADMIN_EMAILS;
      if (Array.isArray(appValue)) {
        return appValue.map(normalizeEmailKeySafe_).filter(Boolean);
      }
      if (typeof appValue === 'string') {
        return appValue
          .split(/[\s,]+/)
          .map(normalizeEmailKeySafe_)
          .filter(Boolean);
      }
    }
  } catch (err) {
    // ignore app config errors
  }

  return [];
}

function assertBillingAdmin_() {
  const email = normalizeEmailKeySafe_(getActiveUserEmail_());
  const admins = getBillingAdminEmails_().map(normalizeEmailKeySafe_).filter(Boolean);
  const isAdmin = email && admins.indexOf(email) !== -1;
  if (!isAdmin) {
    throw new Error('この操作は管理者のみ実行できます');
  }
  return email;
}

function getBillingAdminInfo() {
  const email = normalizeEmailKeySafe_(getActiveUserEmail_());
  const admins = getBillingAdminEmails_().map(normalizeEmailKeySafe_).filter(Boolean);
  const isAdmin = email && admins.indexOf(email) !== -1;
  return { isAdmin, email };
}

function buildPreparedBillingPayload_(billingMonth) {
  if (typeof clearTreatmentLogCache_ === 'function') {
    clearTreatmentLogCache_();
  }
  const resolvedMonthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const source = getBillingSourceData(resolvedMonthKey || billingMonth);
  const existingPrepared = resolvedMonthKey ? loadPreparedBillingFromSheet_(resolvedMonthKey) : null;
  const normalizedExistingPrepared = normalizePreparedBilling_(existingPrepared);
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
  let bankFlagsByPatient = getBankFlagsByPatient_(billingMonthKey || source.billingMonth || billingMonth, {
    patients: patientMap
  });
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
  const payload = {
    schemaVersion: PREPARED_BILLING_SCHEMA_VERSION,
    billingMonth: billingMonthKey || source.billingMonth,
    billingJson: billingJsonArray,
    bankRecords: source.bankRecords || [],
    preparedAt: new Date().toISOString(),
    patients: source.patients || source.patientMap || {},
    bankInfoByName: source.bankInfoByName || {},
    staffByPatient: source.staffByPatient || {},
    staffDirectory: source.staffDirectory || {},
    staffDisplayByPatient: source.staffDisplayByPatient || {},
    billingOverrideFlags: source.billingOverrideFlags || {},
    bankFlagsByPatient,
    carryOverByPatient: source.carryOverByPatient || {},
    carryOverLedger: source.carryOverLedger || [],
    carryOverLedgerMeta: source.carryOverLedgerMeta || {},
    carryOverLedgerByPatient: source.carryOverLedgerByPatient || {},
    unpaidHistory: source.unpaidHistory || [],
    visitsByPatient,
    totalsByPatient,
    bankAccountInfoByPatient,
    receiptStatus: normalizedExistingPrepared && normalizedExistingPrepared.receiptStatus,
    aggregateUntilMonth: normalizedExistingPrepared && normalizedExistingPrepared.aggregateUntilMonth
  };
  return mergeReceiptSettingsIntoPrepared_(payload,
    normalizedExistingPrepared && normalizedExistingPrepared.receiptStatus,
    normalizedExistingPrepared && normalizedExistingPrepared.aggregateUntilMonth);
}

const BANK_WITHDRAWAL_SHEET_PREFIX = '銀行引落_';
const BANK_WITHDRAWAL_AMOUNT_COLUMN_LETTER = 'S';

function generateSimpleBankSheet(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const prepared = prepareBillingData(month.key);

  const workbook = billingSs();
  const templateSheet = ensureBankInfoSheet_();
  const sheetName = formatBankWithdrawalSheetName_(month);

  const existing = workbook.getSheetByName(sheetName);
  if (existing) {
    workbook.deleteSheet(existing);
  }

  const copied = templateSheet.copyTo(workbook);
  copied.setName(sheetName);

  const lastRow = copied.getLastRow();
  const lastCol = copied.getLastColumn();
  if (lastRow < 2) {
    return { billingMonth: month.key, sheetName, rows: 0, filled: 0, missingAccounts: [] };
  }

  const amountColumnIndex = columnLetterToNumber_(BANK_WITHDRAWAL_AMOUNT_COLUMN_LETTER);
  const headerColCount = Math.max(lastCol, amountColumnIndex || 0);
  const initialHeaders = copied.getRange(1, 1, 1, headerColCount).getDisplayValues()[0];
  const pidCol = ensureBankWithdrawalPatientIdColumn_(copied, initialHeaders);
  const refreshedHeaderCount = Math.max(copied.getLastColumn(), headerColCount, pidCol || 0);
  const headers = copied.getRange(1, 1, 1, refreshedHeaderCount).getDisplayValues()[0];
  const nameCol = resolveBillingColumn_(headers, BILLING_LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const kanaCol = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', {});
  const amountCol = resolveBillingColumn_(
    headers,
    ['金額', '請求金額', '引落額', '引落金額'],
    '金額',
    { required: true, fallbackLetter: BANK_WITHDRAWAL_AMOUNT_COLUMN_LETTER }
  );
  const patientIdLabels = (typeof BILLING_LABELS !== 'undefined' && BILLING_LABELS && Array.isArray(BILLING_LABELS.recNo))
    ? BILLING_LABELS.recNo
    : [];
  const resolvedPidCol = pidCol || resolveBillingColumn_(headers, patientIdLabels.concat(['患者ID', '患者番号']), '患者ID', {});

  const rowCount = lastRow - 1;
  const nameValues = copied.getRange(2, nameCol, rowCount, 1).getDisplayValues();
  const kanaValues = kanaCol ? copied.getRange(2, kanaCol, rowCount, 1).getDisplayValues() : [];
  const pidValues = resolvedPidCol ? copied.getRange(2, resolvedPidCol, rowCount, 1).getDisplayValues() : [];
  const patients = resolvePreparedPatients_(prepared);
  const nameToPatientId = buildPatientNameToIdMap_(patients);
  const amountByPatientId = buildBillingAmountByPatientId_(prepared && prepared.billingJson);

  const missingAccounts = [];
  const diagnostics = [];
  const amountValues = nameValues.map((row, idx) => {
    const rawName = row && row[0] ? String(row[0]).trim() : '';
    const rawKana = (kanaValues[idx] && kanaValues[idx][0]) ? String(kanaValues[idx][0]).trim() : '';
    const rawPid = pidValues[idx] && pidValues[idx][0];
    const normalizedPid = typeof billingNormalizePatientId_ === 'function'
      ? billingNormalizePatientId_(rawPid)
      : (rawPid ? String(rawPid).trim() : '');
    const fullNameKey = buildFullNameKey_(rawName, rawKana);
    const pid = normalizedPid || (fullNameKey ? nameToPatientId[fullNameKey] : '');
    const hasAmount = pid && Object.prototype.hasOwnProperty.call(amountByPatientId, pid);
    const amount = hasAmount ? amountByPatientId[pid] : null;
    const resolvedBy = normalizedPid ? 'patientId' : (fullNameKey ? 'name' : 'none');

    if (diagnostics.length < 5) {
      diagnostics.push({ row: idx + 2, rawName, rawKana, rawPid, normalizedPid, fullNameKey, pid, hasAmount, amount, resolvedBy });
    }

    if (!pid || amount === null || amount === undefined) {
      if (rawName) missingAccounts.push(rawName);
      return [''];
    }
    return [amount];
  });

  copied.getRange(2, amountCol, rowCount, 1).setValues(amountValues);
  const filled = amountValues.filter(v => v && v[0] !== '' && v[0] !== null && v[0] !== undefined).length;

  if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
    billingLogger_.log('[billing] generateSimpleBankSheet summary=' + JSON.stringify({
      billingMonth: month.key,
      sheetName,
      rowCount,
      filled,
      missing: missingAccounts.length,
      nameCol,
      kanaCol,
      amountCol,
      pidCol: resolvedPidCol
    }));
    if (filled === 0) {
      billingLogger_.log('[billing] generateSimpleBankSheet diagnostics=' + JSON.stringify(diagnostics));
    }
  }

  return { billingMonth: month.key, sheetName, rows: rowCount, filled, missingAccounts };
}

function ensureBankInfoSheet_() {
  const workbook = billingSs();
  const sheet = workbook && typeof workbook.getSheetByName === 'function'
    ? workbook.getSheetByName(BANK_INFO_SHEET_NAME)
    : null;
  if (sheet) return sheet;

  throw new Error('銀行情報シートが見つかりません。参照専用のテンプレートを用意してください。');
}

function ensureUnpaidCheckColumn_(sheet, headers, options) {
  const ensured = ensureBankWithdrawalFlagColumns_(sheet, headers, options);
  return ensured.unpaidCol;
}

function ensureBankWithdrawalFlagColumns_(sheet, headers, options) {
  const targetSheet = sheet;
  const opts = options || {};
  const workingHeaders = Array.isArray(headers) && headers.length
    ? headers.slice()
    : (targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getDisplayValues()[0] || []);

  let unpaidCol = resolveBillingColumn_(workingHeaders, [BANK_WITHDRAWAL_UNPAID_HEADER], BANK_WITHDRAWAL_UNPAID_HEADER, {});
  if (!unpaidCol) {
    const lastCol = targetSheet.getLastColumn();
    targetSheet.insertColumnAfter(lastCol);
    unpaidCol = lastCol + 1;
    targetSheet.getRange(1, unpaidCol).setValue(BANK_WITHDRAWAL_UNPAID_HEADER);
  }

  let aggregateCol = resolveBillingColumn_(workingHeaders, [BANK_WITHDRAWAL_AGGREGATE_HEADER], BANK_WITHDRAWAL_AGGREGATE_HEADER, {});
  if (!aggregateCol) {
    try {
      targetSheet.insertColumnAfter(unpaidCol);
      aggregateCol = unpaidCol + 1;
      targetSheet.getRange(1, aggregateCol).setValue(BANK_WITHDRAWAL_AGGREGATE_HEADER);
    } catch (err) {
      console.warn('[billing] Failed to insert aggregate column on bank withdrawal sheet', err);
    }
  }

  try {
    const lastRow = targetSheet.getLastRow();
    if (lastRow >= 2) {
      const headerCount = targetSheet.getLastColumn();
      const refreshedHeaders = targetSheet.getRange(1, 1, 1, headerCount).getDisplayValues()[0] || [];
      const rowCount = lastRow - 1;
      const rowValues = targetSheet.getRange(2, 1, rowCount, headerCount).getDisplayValues();
      const hasValue_ = value => value !== '' && value !== null && value !== undefined && String(value).trim() !== '';
      const validRows = [];

      rowValues.forEach((row, index) => {
        if (!row || !row.length) return;
        const hasAnyValue = row.some(value => hasValue_(value));
        if (!hasAnyValue) return;
        validRows.push(index + 2);
      });

      const segments = [];
      if (validRows.length) {
        let segmentStart = validRows[0];
        let segmentLength = 1;
        for (let i = 1; i < validRows.length; i += 1) {
          const rowNumber = validRows[i];
          if (rowNumber === segmentStart + segmentLength) {
            segmentLength += 1;
          } else {
            segments.push({ startRow: segmentStart, length: segmentLength });
            segmentStart = rowNumber;
            segmentLength = 1;
          }
        }
        segments.push({ startRow: segmentStart, length: segmentLength });
      }

      if (segments.length) {
        const rule = SpreadsheetApp && typeof SpreadsheetApp.newDataValidation === 'function'
          ? SpreadsheetApp.newDataValidation().requireCheckbox().build()
          : null;
        const applyCheckboxes = columnIndex => {
          segments.forEach(segment => {
            const range = targetSheet.getRange(segment.startRow, columnIndex, segment.length, 1);
            if (range && typeof range.insertCheckboxes === 'function') {
              range.insertCheckboxes();
            } else if (rule) {
              range.setDataValidation(rule);
            }
          });
        };
        applyCheckboxes(unpaidCol);
        if (aggregateCol) {
          applyCheckboxes(aggregateCol);
        }
      }
    }
  } catch (err) {
    console.warn('[billing] Failed to ensure checkbox columns for bank withdrawal flags', err);
  }

  enforceBankWithdrawalAggregateConstraint_(targetSheet, unpaidCol, aggregateCol);

  return { unpaidCol, aggregateCol };
}

function enforceBankWithdrawalAggregateConstraint_(sheet, unpaidCol, aggregateCol) {
  if (!sheet || !unpaidCol || !aggregateCol) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const rowCount = lastRow - 1;
  const unpaidValues = sheet.getRange(2, unpaidCol, rowCount, 1).getValues();
  const aggregateRange = sheet.getRange(2, aggregateCol, rowCount, 1);
  const aggregateValues = aggregateRange.getValues();

  let needsUpdate = false;
  for (let i = 0; i < rowCount; i++) {
    const unpaid = !!(unpaidValues[i] && unpaidValues[i][0]);
    const aggregate = aggregateValues[i] && aggregateValues[i][0];
    if (!unpaid && aggregate) {
      aggregateValues[i][0] = false;
      needsUpdate = true;
    }
  }

  if (needsUpdate) {
    aggregateRange.setValues(aggregateValues);
  }
}

function formatBankWithdrawalSheetName_(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const monthText = String(month.month).padStart(2, '0');
  return `${BANK_WITHDRAWAL_SHEET_PREFIX}${month.year}-${monthText}`;
}

function refreshBankWithdrawalSheetFromTemplate_(targetSheet, templateSheet) {
  if (!targetSheet || !templateSheet) return;

  const lastRow = templateSheet.getLastRow();
  const lastCol = templateSheet.getLastColumn();
  const templateValues = templateSheet.getRange(1, 1, lastRow, lastCol).getValues();
  targetSheet.getRange(1, 1, lastRow, lastCol).setValues(templateValues);

  const targetLastRow = targetSheet.getLastRow();
  if (targetLastRow > lastRow) {
    const extraRows = targetLastRow - lastRow;
    const clearRange = targetSheet.getRange(lastRow + 1, 1, extraRows, Math.max(targetSheet.getLastColumn(), lastCol));
    if (clearRange && typeof clearRange.clearContent === 'function') {
      clearRange.clearContent();
    } else {
      const blankRows = new Array(extraRows).fill(null).map(() => new Array(lastCol).fill(''));
      clearRange.setValues(blankRows);
    }
  }
}

function ensureBankWithdrawalSheet_(billingMonth, options) {
  const workbook = billingSs();
  const baseSheet = ensureBankInfoSheet_();
  const sheetName = formatBankWithdrawalSheetName_(billingMonth);
  const opts = options || {};
  const shouldRefresh = opts.refreshFromTemplate !== false;
  const preserveExistingSheet = opts.preserveExistingSheet === true;
  const existingSheet = workbook.getSheetByName(sheetName);

  if (existingSheet && shouldRefresh && !preserveExistingSheet) {
    try {
      workbook.deleteSheet(existingSheet);
    } catch (err) {
      console.warn('[billing] Failed to replace existing bank withdrawal sheet, will reuse it', err);
      return existingSheet;
    }
  } else if (existingSheet && shouldRefresh) {
    try {
      refreshBankWithdrawalSheetFromTemplate_(existingSheet, baseSheet);
    } catch (err) {
      console.warn('[billing] Failed to refresh bank withdrawal sheet from template', err);
    }
  }

  let sheet = workbook.getSheetByName(sheetName);
  if (!sheet) {
    sheet = baseSheet.copyTo(workbook);
    sheet.setName(sheetName);
    const targetIndex = Math.max(baseSheet.getIndex() + 1, workbook.getNumSheets());
    workbook.setActiveSheet(sheet);
    workbook.moveActiveSheet(targetIndex);
    try {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET) || [];
      protections.forEach(protection => protection.remove());
    } catch (err) {
      console.warn('[billing] Failed to remove protection from bank withdrawal sheet', err);
    }
  }
  ensureBankWithdrawalPatientIdColumn_(sheet);
  ensureUnpaidCheckColumn_(sheet, null, {
    billingJson: opts.billingJson,
    bankRecords: opts.bankRecords
  });
  return sheet;
}

function ensureBankWithdrawalPatientIdColumn_(sheet, headers) {
  const targetSheet = sheet;
  const workingHeaders = Array.isArray(headers) && headers.length
    ? headers.slice()
    : (targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getDisplayValues()[0] || []);
  const patientIdLabels = (typeof BILLING_LABELS !== 'undefined' && BILLING_LABELS && Array.isArray(BILLING_LABELS.recNo))
    ? BILLING_LABELS.recNo
    : [];
  let col = resolveBillingColumn_(workingHeaders, patientIdLabels.concat(['患者ID', '患者番号']), '患者ID', {});
  if (col) return col;

  const nameCol = resolveBillingColumn_(workingHeaders, BILLING_LABELS.name, '名前', { fallbackIndex: 1 });
  const kanaCol = resolveBillingColumn_(workingHeaders, BILLING_LABELS.furigana, 'フリガナ', {});
  const insertAfter = kanaCol || nameCol || Math.max(targetSheet.getLastColumn(), workingHeaders.length);

  try {
    if (typeof targetSheet.insertColumnAfter === 'function') {
      targetSheet.insertColumnAfter(insertAfter);
    } else if (typeof targetSheet.insertColumnBefore === 'function') {
      targetSheet.insertColumnBefore(insertAfter + 1);
    }
    col = insertAfter + 1;
    if (typeof targetSheet.getRange === 'function') {
      const headerCell = targetSheet.getRange(1, col, 1, 1);
      if (headerCell && typeof headerCell.setValue === 'function') {
        headerCell.setValue('患者ID');
      } else if (headerCell && typeof headerCell.setValues === 'function') {
        headerCell.setValues([['患者ID']]);
      }
    }
  } catch (err) {
    console.warn('[billing] Failed to ensure patient ID column on bank withdrawal sheet', err);
  }

  return col;
}

function normalizeBillingNameKey_(value) {
  return String(value || '')
    .normalize('NFKC')
    .replace(/[\s\u3000・･．.ー－ｰ−‐-]+/g, '')
    .trim();
}

function buildFullNameKey_(nameKanji, nameKana) {
  const kanjiKey = normalizeBillingNameKey_(nameKanji);
  const kanaKey = normalizeBillingNameKey_(nameKana);
  const combined = [kanjiKey, kanaKey].filter(Boolean).join('::');
  if (!combined) return '';
  const numericOnly = combined.replace(/::/g, '');
  if (/^\d+$/.test(numericOnly)) return '';
  return combined;
}

function buildPatientNameToIdMap_(patients) {
  const entries = patients && typeof patients === 'object' ? Object.keys(patients) : [];
  return entries.reduce((map, pid) => {
    const patient = patients[pid] || {};
    const key = buildFullNameKey_(
      patient.nameKanji || (patient.raw && (patient.raw.nameKanji || patient.raw['氏名'])),
      patient.nameKana || (patient.raw && (patient.raw.nameKana || patient.raw['フリガナ']))
    );
    const normalizedPid = billingNormalizePatientId_(pid);
    if (key && normalizedPid && !map[key]) {
      map[key] = normalizedPid;
    }
    return map;
  }, {});
}

function resolvePreparedPatients_(preparedOrPatients) {
  if (!preparedOrPatients || typeof preparedOrPatients !== 'object') return {};
  if (preparedOrPatients.patients && typeof preparedOrPatients.patients === 'object') {
    return preparedOrPatients.patients || {};
  }
  return preparedOrPatients;
}

function normalizeBankFlagValue_(value) {
  if (value === true) return true;
  if (value === false || value === null || value === undefined) return false;
  const num = Number(value);
  if (Number.isFinite(num)) return num !== 0;
  const normalized = String(value || '').trim();
  if (!normalized) return false;
  const lowered = normalized.toLowerCase();
  if (lowered === 'true' || lowered === 'yes' || lowered === 'on') return true;
  if (lowered === 'false' || lowered === 'no' || lowered === 'off') return false;
  return lowered === '1' || normalized === '✓' || normalized === '✔' || normalized === '☑' || normalized === '◯';
}

function getBankFlagsByPatient_(billingMonth, prepared) {
  const month = normalizeBillingMonthInput(billingMonth);
  const workbook = billingSs();
  const sheetName = formatBankWithdrawalSheetName_(month);
  const sheet = workbook.getSheetByName(sheetName);
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const unpaidCol = resolveBillingColumn_(headers, [BANK_WITHDRAWAL_UNPAID_HEADER], BANK_WITHDRAWAL_UNPAID_HEADER, {});
  const aggregateCol = resolveBillingColumn_(headers, ['合算'], '合算', {});
  const nameCol = resolveBillingColumn_(headers, BILLING_LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const kanaCol = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', {});
  const pidCol = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', {});

  const effectiveLastCol = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, effectiveLastCol).getValues();
  const patients = resolvePreparedPatients_(prepared);
  const nameToPatientId = buildPatientNameToIdMap_(patients);
  const flagsByPatient = {};

  values.forEach(row => {
    const rawPid = pidCol ? row[pidCol - 1] : '';
    const normalizedPid = billingNormalizePatientId_(rawPid);
    const pid = normalizedPid || (nameCol
      ? nameToPatientId[buildFullNameKey_(row[nameCol - 1], kanaCol ? row[kanaCol - 1] : '')]
      : '');
    if (!pid) return;

    const current = flagsByPatient[pid] || { ae: false, af: false };
    const ae = unpaidCol ? normalizeBankFlagValue_(row[unpaidCol - 1]) : false;
    const af = aggregateCol ? normalizeBankFlagValue_(row[aggregateCol - 1]) : false;
    flagsByPatient[pid] = { ae: current.ae || ae, af: current.af || af };
  });

  return flagsByPatient;
}

function buildBillingAmountByPatientId_(billingJson) {
  const amounts = {};
  (billingJson || []).forEach(entry => {
    const pid = billingNormalizePatientId_(entry && entry.patientId);
    if (!pid) return;
    const amountCandidate = entry && entry.grandTotal != null
      ? entry.grandTotal
      : (entry && entry.total != null ? entry.total : entry && entry.billingAmount);
    const amount = typeof normalizeMoneyNumber_ === 'function'
      ? normalizeMoneyNumber_(amountCandidate)
      : Number(amountCandidate) || 0;
    amounts[pid] = amount;
  });
  return amounts;
}

function syncBankWithdrawalSheetForMonth_(billingMonth, prepared) {
  const month = normalizeBillingMonthInput(billingMonth || (prepared && prepared.billingMonth));
  const sheet = ensureBankWithdrawalSheet_(month, {
    refreshFromTemplate: true,
    preserveExistingSheet: true,
    billingJson: prepared && prepared.billingJson,
    bankRecords: prepared && prepared.bankRecords
  });
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { billingMonth: month.key, updated: 0 };

  const lastCol = sheet.getLastColumn();
  const initialHeaders = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const pidCol = ensureBankWithdrawalPatientIdColumn_(sheet, initialHeaders);
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), lastCol, pidCol || 0)).getDisplayValues()[0];
  const nameCol = resolveBillingColumn_(headers, BILLING_LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const kanaCol = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', {});
  const amountCol = resolveBillingColumn_(
    headers,
    ['金額', '請求金額', '引落額', '引落金額'],
    '金額',
    { required: true, fallbackLetter: BANK_WITHDRAWAL_AMOUNT_COLUMN_LETTER }
  );
  const patientIdLabels = (typeof BILLING_LABELS !== 'undefined' && BILLING_LABELS && Array.isArray(BILLING_LABELS.recNo))
    ? BILLING_LABELS.recNo
    : [];
  const resolvedPidCol = pidCol || resolveBillingColumn_(headers, patientIdLabels.concat(['患者ID', '患者番号']), '患者ID', {});

  const rowCount = lastRow - 1;
  const nameValues = sheet.getRange(2, nameCol, rowCount, 1).getDisplayValues();
  const kanaValues = kanaCol ? sheet.getRange(2, kanaCol, rowCount, 1).getDisplayValues() : [];
  const pidValues = resolvedPidCol ? sheet.getRange(2, resolvedPidCol, rowCount, 1).getDisplayValues() : [];
  const existingAmountValues = sheet.getRange(2, amountCol, rowCount, 1).getValues();
  const patients = resolvePreparedPatients_(prepared);
  const nameToPatientId = buildPatientNameToIdMap_(patients);
  const amountByPatientId = buildBillingAmountByPatientId_(prepared && prepared.billingJson);
  const isBlank_ = value => value === '' || value === null || value === undefined;
  const isSameAmount_ = (current, next) => {
    if (isBlank_(current) && isBlank_(next)) return true;
    const currentNum = !isBlank_(current) ? Number(current) : NaN;
    const nextNum = !isBlank_(next) ? Number(next) : NaN;
    if (Number.isFinite(currentNum) && Number.isFinite(nextNum)) {
      return currentNum === nextNum;
    }
    return String(current) === String(next);
  };

  const newAmountValues = nameValues.map((row, idx) => {
    const kanaRow = kanaValues[idx] || [];
    const rawPid = pidValues[idx] && pidValues[idx][0];
    const normalizedPid = typeof billingNormalizePatientId_ === 'function'
      ? billingNormalizePatientId_(rawPid)
      : (rawPid ? String(rawPid).trim() : '');
    const nameKey = buildFullNameKey_(row && row[0], kanaRow[0]);
    const pid = normalizedPid || (nameKey ? nameToPatientId[nameKey] : '');
    const resolvedAmount = pid && Object.prototype.hasOwnProperty.call(amountByPatientId, pid)
      ? amountByPatientId[pid]
      : null;
    const existingAmount = existingAmountValues[idx][0];
    const hasManualAmount = existingAmount !== '' && existingAmount !== null && existingAmount !== undefined;
    const nextAmount = hasManualAmount && resolvedAmount !== null && resolvedAmount !== undefined
      ? existingAmount
      : (resolvedAmount !== null && resolvedAmount !== undefined ? resolvedAmount : existingAmount);
    return [nextAmount];
  });

  const updates = [];
  newAmountValues.forEach((rowValue, idx) => {
    const existingValue = existingAmountValues[idx] ? existingAmountValues[idx][0] : '';
    const nextValue = rowValue[0];
    if (!isSameAmount_(existingValue, nextValue)) {
      updates.push({ row: idx + 2, value: nextValue });
    }
  });

  if (!updates.length) {
    return { billingMonth: month.key, updated: 0 };
  }

  let segmentStart = updates[0].row;
  let segmentValues = [[updates[0].value]];
  let previousRow = updates[0].row;
  for (let idx = 1; idx < updates.length; idx++) {
    const update = updates[idx];
    if (update.row === previousRow + 1) {
      segmentValues.push([update.value]);
    } else {
      sheet.getRange(segmentStart, amountCol, segmentValues.length, 1).setValues(segmentValues);
      segmentStart = update.row;
      segmentValues = [[update.value]];
    }
    previousRow = update.row;
  }
  sheet.getRange(segmentStart, amountCol, segmentValues.length, 1).setValues(segmentValues);

  return { billingMonth: month.key, updated: updates.length };
}

function summarizeBankWithdrawalSheet_(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const workbook = billingSs();
  const sheetName = formatBankWithdrawalSheetName_(month);
  const sheet = workbook.getSheetByName(sheetName);
  if (!sheet) {
    return { billingMonth: month.key, exists: false, rows: 0, unpaidChecked: 0 };
  }

  const lastRow = sheet.getLastRow();
  const rows = Math.max(lastRow - 1, 0);
  if (!rows) {
    return { billingMonth: month.key, exists: true, rows: 0, unpaidChecked: 0 };
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const unpaidCol = ensureUnpaidCheckColumn_(sheet, headers);
  const checkedRows = sheet.getRange(2, unpaidCol, rows, 1).getValues().filter(row => row[0]).length;

  return { billingMonth: month.key, exists: true, rows, unpaidChecked: checkedRows };
}

function collectUnpaidWithdrawalRows_(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const workbook = billingSs();
  const sheetName = formatBankWithdrawalSheetName_(month);
  const sheet = workbook.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('銀行引落シートが見つかりません: ' + sheetName);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { billingMonth: month.key, entries: [], checkedRows: 0 };

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const unpaidCol = ensureUnpaidCheckColumn_(sheet, headers);
  const amountCol = resolveBillingColumn_(
    headers,
    ['金額', '請求金額', '引落額', '引落金額'],
    '金額',
    { required: true, fallbackLetter: BANK_WITHDRAWAL_AMOUNT_COLUMN_LETTER }
  );
  const reasonCol = resolveBillingColumn_(headers, ['未回収理由', '未回収メモ', '未回収備考', '理由'], '未回収理由', {});
  const memoCol = resolveBillingColumn_(headers, ['備考', 'メモ', 'コメント', '未回収備考'], '備考', {});
  const pidCol = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', {});
  const nameCol = resolveBillingColumn_(headers, BILLING_LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const kanaCol = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', {});
  const effectiveLastCol = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, effectiveLastCol).getValues();
  const prepared = loadPreparedBillingWithSheetFallback_(month.key, { withValidation: true }) || {};
  const patients = resolvePreparedPatients_(prepared);
  const nameToPatientId = buildPatientNameToIdMap_(patients);
  const entries = [];
  let checkedRows = 0;

  values.forEach(row => {
    const isChecked = row[unpaidCol - 1];
    if (!isChecked) return;
    checkedRows += 1;
    const amount = amountCol ? Number(row[amountCol - 1]) || 0 : 0;
    if (!amount) return;
    const reason = reasonCol ? String(row[reasonCol - 1] || '').trim() : '';
    const memo = memoCol ? String(row[memoCol - 1] || '').trim() : '';
    const pid = pidCol
      ? billingNormalizePatientId_(row[pidCol - 1])
      : nameToPatientId[buildFullNameKey_(row[nameCol - 1], kanaCol ? row[kanaCol - 1] : '')];
    if (!pid) return;
    entries.push({
      patientId: pid,
      billingMonth: month.key,
      unpaidAmount: amount,
      reason: reason || BANK_WITHDRAWAL_UNPAID_HEADER,
      memo
    });
  });

  return { billingMonth: month.key, entries, checkedRows };
}

function applyBankWithdrawalUnpaidEntries(billingMonth) {
  const collected = collectUnpaidWithdrawalRows_(billingMonth);
  if (!collected.entries.length) {
    return Object.assign({ added: 0, skipped: 0 }, collected);
  }
  const appendResult = appendUnpaidHistoryEntries_(collected.entries);
  return Object.assign({}, collected, appendResult);
}

function prepareBankWithdrawalData(billingMonth) {
  const prepared = prepareBillingData(billingMonth);
  return summarizeBankWithdrawalState_(billingMonth, prepared, summarizeBankWithdrawalSheet_(billingMonth), {
    phase: 'aggregated'
  });
}

function generateBankWithdrawalSheetFromCache(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const loaded = loadPreparedBillingWithSheetFallback_(month.key, { withValidation: true, restoreCache: true });
  const validation = loaded && loaded.validation ? loaded.validation : null;
  const prepared = normalizePreparedBilling_(loaded && loaded.prepared);

  if (!prepared || !prepared.billingJson || (validation && validation.ok === false)) {
    throw new Error('請求データが未集計です。先に「請求データ集計」を実行してください。');
  }

  const bankSheet = syncBankWithdrawalSheetForMonth_(month, prepared);
  const sheetSummary = summarizeBankWithdrawalSheet_(month);

  return summarizeBankWithdrawalState_(month, prepared, sheetSummary, {
    bankSheet,
    validation
  });
}

function applyBankWithdrawalUnpaidFromUi(billingMonth) {
  const result = applyBankWithdrawalUnpaidEntries(billingMonth);
  const summary = summarizeBankWithdrawalSheet_(billingMonth);
  return Object.assign({}, result, {
    billingMonth: normalizeBillingMonthInput(billingMonth).key,
    sheetSummary: summary
  });
}

function summarizeBankWithdrawalState_(billingMonth, prepared, sheetSummary, extra) {
  const normalizedMonth = normalizeBillingMonthInput(billingMonth);
  const normalizedPrepared = normalizePreparedBilling_(prepared);
  const billingCount = normalizedPrepared && Array.isArray(normalizedPrepared.billingJson)
    ? normalizedPrepared.billingJson.length
    : 0;
  const summary = sheetSummary || summarizeBankWithdrawalSheet_(normalizedMonth);

  return Object.assign({
    billingMonth: normalizedMonth.key,
    billingCount,
    preparedAt: normalizedPrepared && normalizedPrepared.preparedAt || null,
    sheetSummary: summary
  }, extra || {});
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
    const normalizeBankFlags_ = flags => ({
      ae: !!(flags && flags.ae),
      af: !!(flags && flags.af)
    });
    const bankFlagsByPatient = normalizeMap_(payload.bankFlagsByPatient);
    const billingJson = coerceBillingJsonArray_(payload.billingJson).map(entry => {
      const pid = typeof billingNormalizePatientId_ === 'function'
        ? billingNormalizePatientId_(entry && entry.patientId)
        : String(entry && entry.patientId ? entry.patientId : '').trim();
      const bankFlags = pid && bankFlagsByPatient && Object.prototype.hasOwnProperty.call(bankFlagsByPatient, pid)
        ? bankFlagsByPatient[pid]
        : null;
      const normalizedEntry = Object.assign({}, entry || {}, {
        bankFlags: normalizeBankFlags_(bankFlags)
      });
      return sanitizeAggregateFieldsForBankFlags_(normalizedEntry, normalizedEntry.bankFlags);
    });
    const schemaVersion = Number(payload.schemaVersion);
    const normalized = {
      schemaVersion: Number.isFinite(schemaVersion) ? schemaVersion : null,
      billingMonth: payload.billingMonth || '',
      preparedAt: payload.preparedAt || null,
      billingJson,
      patients: normalizeMap_(payload.patients),
      bankRecords: Array.isArray(payload.bankRecords) ? payload.bankRecords : [],
      bankInfoByName: normalizeMap_(payload.bankInfoByName),
      bankAccountInfoByPatient: normalizeMap_(payload.bankAccountInfoByPatient),
      bankFlagsByPatient,
      visitsByPatient: normalizeMap_(payload.visitsByPatient),
      totalsByPatient: normalizeMap_(payload.totalsByPatient),
      staffByPatient: normalizeMap_(payload.staffByPatient),
      staffDirectory: normalizeMap_(payload.staffDirectory),
      staffDisplayByPatient: normalizeMap_(payload.staffDisplayByPatient),
      billingOverrideFlags: normalizeMap_(payload.billingOverrideFlags),
      bankFlagsByPatient: normalizeMap_(payload.bankFlagsByPatient),
      carryOverByPatient: normalizeMap_(payload.carryOverByPatient),
      carryOverLedger: Array.isArray(payload.carryOverLedger) ? payload.carryOverLedger : [],
      carryOverLedgerMeta: normalizeMap_(payload.carryOverLedgerMeta),
      carryOverLedgerByPatient: normalizeMap_(payload.carryOverLedgerByPatient),
      unpaidHistory: Array.isArray(payload.unpaidHistory) ? payload.unpaidHistory : [],
      bankStatuses: normalizeMap_(payload.bankStatuses),
      receiptStatus: normalizeReceiptStatus_(payload.receiptStatus),
      aggregateUntilMonth: normalizeBillingMonthKeySafe_(payload.aggregateUntilMonth)
    };
    const normalizedLength = Array.isArray(normalized.billingJson) ? normalized.billingJson.length : 0;
    if (isBillingDebugEnabled_()) {
      billingLogger_.log('[billing] normalizePreparedBilling_ lengths=' + JSON.stringify({
        rawBillingJsonLength: rawLength,
        normalizedBillingJsonLength: normalizedLength,
        billingMonth: normalized.billingMonth
      }));
    }
    return normalized;
  }

function toClientBillingPayload_(prepared, options) {
  const rawLength = prepared && prepared.billingJson
    ? (Array.isArray(prepared.billingJson) ? prepared.billingJson.length : 'non-array')
    : 0;
  const opts = options || {};
  const normalized = opts.alreadyNormalized ? prepared : normalizePreparedBilling_(prepared);
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
    bankFlagsByPatient: normalized.bankFlagsByPatient || {},
    carryOverByPatient: normalized.carryOverByPatient || {},
    carryOverLedger: normalized.carryOverLedger || [],
    carryOverLedgerMeta: normalized.carryOverLedgerMeta || {},
    carryOverLedgerByPatient: normalized.carryOverLedgerByPatient || {},
    unpaidHistory: normalized.unpaidHistory || [],
    visitsByPatient: normalized.visitsByPatient || {},
    totalsByPatient: normalized.totalsByPatient || {},
    bankAccountInfoByPatient: normalized.bankAccountInfoByPatient || {},
    bankStatuses: normalized.bankStatuses || {},
    receiptStatus: normalized.receiptStatus || '',
    aggregateUntilMonth: normalized.aggregateUntilMonth || ''
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

function deletePreparedBillingRowsForMonth_(sheet, monthKey) {
  const sheetName = sheet && typeof sheet.getName === 'function' ? sheet.getName() : '';
  if (!sheet || !monthKey) return { sheetName, removed: 0 };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { sheetName, removed: 0 };

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let removed = 0;
  for (let idx = values.length - 1; idx >= 0; idx--) {
    if (String(values[idx][0] || '').trim() !== monthKey) continue;
    sheet.deleteRow(idx + 2);
    removed += 1;
  }

  return { sheetName, removed };
}

function deletePreparedBillingDataForMonth_(billingMonth) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const summary = { billingMonth: monthKey || '', cacheCleared: false, sheetDeletions: [] };
  if (!monthKey) return summary;

  const cacheKey = buildBillingCacheKey_(monthKey);
  if (cacheKey) {
    clearBillingCache_(cacheKey);
    summary.cacheCleared = true;
  }

  const workbook = billingSs();
  if (workbook) {
    const removeRows = sheetName => deletePreparedBillingRowsForMonth_(workbook.getSheetByName(sheetName), monthKey);
    summary.sheetDeletions = [
      removeRows('PreparedBillingMeta'),
      removeRows('PreparedBillingMetaJson'),
      removeRows('PreparedBillingJson')
    ];
  }

  try {
    if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
      billingLogger_.log('[billing] deletePreparedBillingDataForMonth_ summary=' + JSON.stringify(summary));
    }
  } catch (err) {
    // ignore logging issues in non-GAS environments
  }

  return summary;
}

function resetPreparedBillingAndPrepare(billingMonth) {
  const normalizedMonth = normalizeBillingMonthInput ? normalizeBillingMonthInput(billingMonth) : null;
  const monthKey = normalizedMonth && normalizedMonth.key
    ? normalizedMonth.key
    : normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) {
    throw new Error('請求月をYYYY-MM形式で指定してください。');
  }

  deletePreparedBillingDataForMonth_(monthKey);
  try {
    if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
      billingLogger_.log('[billing] resetPreparedBillingAndPrepare clearing existing data for ' + monthKey);
    }
  } catch (err) {
    // ignore logging issues in non-GAS environments
  }

  return prepareBillingData(monthKey);
}

function prepareBillingData(billingMonth) {
  const normalizedMonth = normalizeBillingMonthInput(billingMonth);
  const existingResult = loadPreparedBillingWithSheetFallback_(normalizedMonth.key, { withValidation: true });
  const existingPrepared = existingResult && existingResult.prepared !== undefined ? existingResult.prepared : existingResult;
  const existingValidation = existingResult && existingResult.validation
    ? existingResult.validation
    : (existingResult && existingResult.ok !== undefined ? existingResult : null);
  if (existingPrepared && (!existingValidation || existingValidation.ok)) {
    billingLogger_.log('[billing] prepareBillingData using existing prepared billing for ' + normalizedMonth.key);
    return existingPrepared;
  }

  const prepared = buildPreparedBillingPayload_(normalizedMonth);
  const normalizedPrepared = normalizePreparedBilling_(prepared);
  const clientPayload = toClientBillingPayload_(normalizedPrepared, { alreadyNormalized: true });
  const payloadWithMonth = clientPayload
    ? Object.assign({}, clientPayload, { billingMonth: normalizedMonth.key })
    : clientPayload;
  const serialized = serializeBillingPayload_(payloadWithMonth) || payloadWithMonth;
  const payloadJson = serialized ? JSON.stringify(serialized) : '';

  billingLogger_.log(
    '[billing] prepareBillingData payloadSummary=' +
      JSON.stringify({
        billingMonth: serialized && serialized.billingMonth,
        preparedAt: serialized && serialized.preparedAt,
        billingJsonLength: serialized && serialized.billingJson ? serialized.billingJson.length : null,
        patientCount: serialized && serialized.patients ? Object.keys(serialized.patients).length : null,
        payloadByteLength: payloadJson.length
      })
  );

  const cachePayload = serialized
    ? Object.assign({}, serialized, { billingMonth: normalizedMonth.key })
    : { billingMonth: normalizedMonth.key };
  savePreparedBilling_(cachePayload);
  savePreparedBillingToSheet_(normalizedMonth.key, cachePayload);
  const bankSheetResult = syncBankWithdrawalSheetForMonth_(normalizedMonth, prepared);
  billingLogger_.log('[billing] Bank withdrawal sheet synced: ' + JSON.stringify(bankSheetResult));
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

function normalizeInvoicePatientIdsForGeneration_(patientIds) {
  const source = Array.isArray(patientIds) ? patientIds : String(patientIds || '').split(/[,\s、]+/);
  const normalized = source
    .map(id => String(id || '').trim())
    .filter(Boolean);
  const seen = new Set();
  const unique = [];
  normalized.forEach(id => {
    if (!seen.has(id)) {
      seen.add(id);
      unique.push(id);
    }
  });
  return unique;
}

function filterBillingJsonForInvoice_(billingJson, patientIds) {
  if (!Array.isArray(billingJson)) return [];
  const targets = normalizeInvoicePatientIdsForGeneration_(patientIds);
  if (!targets.length) return billingJson;
  const targetSet = new Set(targets.map(id => String(id).trim()));
  return billingJson.filter(item => targetSet.has(String(item && item.patientId ? item.patientId : '').trim()));
}

function buildStandardInvoiceAmountDataForPdf_(entry, billingMonth) {
  const targetMonth = normalizeBillingMonthKeySafe_(billingMonth || (entry && entry.billingMonth));
  const breakdown = calculateInvoiceChargeBreakdown_(Object.assign({}, entry, { billingMonth: targetMonth }));
  const visits = breakdown.visits || 0;
  const carryOverAmount = normalizeBillingCarryOver_(entry);
  const unitPrice = breakdown.treatmentUnitPrice || 0;
  const transportDetail = breakdown.transportDetail || (formatBillingCurrency_(TRANSPORT_PRICE) + '円 × ' + visits + '回');
  const rows = [
    { label: '前月繰越', detail: '', amount: carryOverAmount },
    { label: '施術料', detail: formatBillingCurrency_(unitPrice) + '円 × ' + visits + '回', amount: breakdown.treatmentAmount || 0 },
    { label: '交通費', detail: transportDetail, amount: breakdown.transportAmount || 0 }
  ];

  return {
    rows,
    grandTotal: breakdown.grandTotal || 0
  };
}

function buildAggregateInvoiceAmountDataForPdf_(aggregateMonths, billingMonth, patientId, cache) {
  const pid = billingNormalizePatientId_(patientId);
  const months = normalizePastBillingMonths_(aggregateMonths, billingMonth);
  const aggregateMonthTotals = months.map(month => ({
    month,
    monthLabel: normalizeBillingMonthLabel_(month),
    total: resolveBillingAmountForMonthAndPatient_(month, pid, null, cache)
  }));
  const grandTotal = aggregateMonthTotals.reduce((sum, row) => sum + (Number(row.total) || 0), 0);
  return {
    aggregateMonthTotals,
    aggregateRemark: formatAggregateBillingRemark_(months),
    grandTotal
  };
}

function finalizeInvoiceAmountDataForPdf_(entry, billingMonth, aggregateMonths, isAggregateInvoice, cache) {
  const normalizedAggregateMonths = normalizePastBillingMonths_(aggregateMonths, billingMonth);
  const baseAmount = isAggregateInvoice
    ? buildAggregateInvoiceAmountDataForPdf_(normalizedAggregateMonths, billingMonth, entry && entry.patientId, cache)
    : buildStandardInvoiceAmountDataForPdf_(entry, billingMonth);
  if (entry && entry.previousReceiptAmount == null) {
    if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
      billingLogger_.log('[billing] finalizeInvoiceAmountDataForPdf_ missing previousReceiptAmount: ' + JSON.stringify({
        patientId: entry.patientId,
        billingMonth,
        hasReceiptMonthBreakdown: Array.isArray(entry.receiptMonthBreakdown),
        receiptMonthBreakdownLength: Array.isArray(entry.receiptMonthBreakdown) ? entry.receiptMonthBreakdown.length : 0
      }));
    }
  }
  const amount = Object.assign({}, baseAmount, {
    insuranceType: entry && entry.insuranceType ? String(entry.insuranceType).trim() : '',
    burdenRate: entry && entry.burdenRate != null ? entry.burdenRate : '',
    chargeMonthLabel: normalizeBillingMonthLabel_(billingMonth),
    forceHideReceipt: !!(entry && entry.forceHideReceipt)
  });

  const receiptDisplay = resolveInvoiceReceiptDisplay_(entry, { aggregateMonths: normalizedAggregateMonths });
  const basePreviousReceipt = buildInvoicePreviousReceipt_(entry, receiptDisplay, normalizedAggregateMonths);
  const previousReceipt = entry && entry.previousReceipt
    ? Object.assign({}, basePreviousReceipt, entry.previousReceipt)
    : basePreviousReceipt;
  if (previousReceipt) {
    previousReceipt.settled = isPreviousReceiptSettled_(entry);
    previousReceipt.visible = !!(receiptDisplay && receiptDisplay.visible);
  }

  const aggregateStatus = receiptDisplay ? receiptDisplay.aggregateStatus : normalizeAggregateStatus_(entry && entry.aggregateStatus);
  const aggregateConfirmed = !!(entry && entry.bankFlags && entry.bankFlags.af === true);
  const watermark = buildInvoiceWatermark_(entry);
  const receiptMonths = receiptDisplay && receiptDisplay.receiptMonths ? receiptDisplay.receiptMonths : [];
  logReceiptDebug_(entry && entry.patientId, {
    step: 'finalizeInvoiceAmountDataForPdf_',
    billingMonth,
    patientId: entry && entry.patientId,
    currentFlags: entry && entry.bankFlags,
    previousFlags: entry && entry.previousBankFlags,
    receiptTargetMonths: normalizedAggregateMonths,
    receiptMonths
  });

  return Object.assign({}, amount, {
    aggregateStatus,
    aggregateConfirmed,
    receiptMonths,
    receiptRemark: receiptDisplay && receiptDisplay.receiptRemark ? receiptDisplay.receiptRemark : '',
    receiptMonthBreakdown: Array.isArray(entry && entry.receiptMonthBreakdown) ? entry.receiptMonthBreakdown : undefined,
    previousReceiptAmount: entry && entry.previousReceiptAmount != null ? entry.previousReceiptAmount : undefined,
    showReceipt: !!(receiptDisplay && receiptDisplay.visible),
    previousReceipt,
    watermark
  });
}

function buildInvoicePdfContextForEntry_(entry, prepared, cache) {
  if (!entry) return null;
  const billingMonth = normalizeBillingMonthKeySafe_(prepared && prepared.billingMonth ? prepared.billingMonth : entry.billingMonth);
  const patientId = billingNormalizePatientId_(entry && entry.patientId);
  if (!billingMonth || !patientId) return null;

  const ensurePreviousReceiptAmount = () => {
    if (entry && entry.previousReceiptAmount != null) return entry;
    if (entry && Array.isArray(entry.receiptMonthBreakdown) && entry.receiptMonthBreakdown.length) {
      const hasBreakdownAmount = entry.receiptMonthBreakdown.some(item => item && item.amount != null && item.amount !== '');
      if (hasBreakdownAmount) {
        const breakdownAmount = entry.receiptMonthBreakdown.reduce((sum, item) => {
          const normalized = normalizeMoneyNumber_(item && item.amount);
          return normalized != null ? sum + normalized : sum;
        }, 0);
        if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
          billingLogger_.log('[billing] buildInvoicePdfContextForEntry_ filled previousReceiptAmount from receiptMonthBreakdown: ' + JSON.stringify({
            patientId,
            billingMonth,
            breakdownAmount
          }));
        }
        return Object.assign({}, entry, { previousReceiptAmount: breakdownAmount });
      }
      return entry;
    }
    const previousMonthKey = resolvePreviousBillingMonthKey_(billingMonth);
    if (!previousMonthKey) return entry;
    const monthCache = cache || {
      preparedByMonth: {},
      bankWithdrawalUnpaidByMonth: {},
      bankWithdrawalAmountsByMonth: {}
    };
    const previousPrepared = getPreparedBillingForMonthCached_(previousMonthKey, monthCache);
    const previousEntry = getPreparedBillingEntryForPatient_(previousPrepared, patientId);
    if (!previousEntry) return entry;
    const previousAmount = previousEntry.grandTotal != null && previousEntry.grandTotal !== ''
      ? normalizeMoneyNumber_(previousEntry.grandTotal)
      : normalizeMoneyNumber_(previousEntry.billingAmount);
    if (previousAmount == null) return entry;
    if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
      billingLogger_.log('[billing] buildInvoicePdfContextForEntry_ filled previousReceiptAmount from previous month: ' + JSON.stringify({
        patientId,
        billingMonth,
        previousMonthKey,
        previousAmount
      }));
    }
    return Object.assign({}, entry, { previousReceiptAmount: previousAmount });
  };

  const receiptEntry = ensurePreviousReceiptAmount();
  if (receiptEntry && receiptEntry.previousReceiptAmount == null) {
    if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
      billingLogger_.log('[billing] buildInvoicePdfContextForEntry_ previousReceiptAmount still missing before finalize: ' + JSON.stringify({
        patientId,
        billingMonth,
        hasReceiptMonthBreakdown: Array.isArray(receiptEntry.receiptMonthBreakdown),
        receiptMonthBreakdownLength: Array.isArray(receiptEntry.receiptMonthBreakdown) ? receiptEntry.receiptMonthBreakdown.length : 0
      }));
    }
  }
  const decision = resolveInvoiceModeFromBankFlags_(billingMonth, patientId, cache);
  const decisionMonths = decision && Array.isArray(decision.months) ? decision.months : [];
  const aggregateMonths = normalizePastBillingMonths_(decisionMonths, billingMonth);
  const isAggregateInvoice = !!(decision && decision.mode === 'aggregate' && aggregateMonths.length > 1);
  const amount = finalizeInvoiceAmountDataForPdf_(receiptEntry, billingMonth, aggregateMonths, isAggregateInvoice, cache);

  return {
    patientId,
    billingMonth,
    months: isAggregateInvoice ? aggregateMonths : [],
    amount,
    name: receiptEntry && receiptEntry.nameKanji ? String(receiptEntry.nameKanji) : '',
    isAggregateInvoice,
    responsibleName: receiptEntry && receiptEntry.responsibleName ? String(receiptEntry.responsibleName) : ''
  };
}

function buildAggregateInvoicePdfContext_(entry, aggregateMonths, prepared, cache) {
  if (!entry) return null;
  const billingMonth = normalizeBillingMonthKeySafe_(prepared && prepared.billingMonth ? prepared.billingMonth : entry.billingMonth);
  const patientId = billingNormalizePatientId_(entry && entry.patientId);
  if (!billingMonth || !patientId) return null;

  const normalizedMonths = normalizePastBillingMonths_(aggregateMonths, billingMonth);
  const amount = finalizeInvoiceAmountDataForPdf_(entry, billingMonth, normalizedMonths, true, cache);

  return {
    patientId,
    billingMonth,
    months: normalizedMonths,
    amount,
    name: entry && entry.nameKanji ? String(entry.nameKanji) : '',
    isAggregateInvoice: true,
    responsibleName: entry && entry.responsibleName ? String(entry.responsibleName) : ''
  };
}

function generatePreparedInvoices_(prepared, options) {
  const normalized = normalizePreparedBilling_(prepared);
  const monthCache = createBillingMonthCache_();
  const aggregateApplied = applyAggregateInvoiceRulesFromBankFlags_(normalized, monthCache);
  const receiptEnriched = attachPreviousReceiptAmounts_(aggregateApplied, monthCache);
  if (!receiptEnriched || !receiptEnriched.billingJson) {
    throw new Error('請求集計結果が見つかりません。先に集計を実行してください。');
  }
  const targetPatientIds = normalizeInvoicePatientIdsForGeneration_(options && options.invoicePatientIds);
  const targetBillingRows = filterBillingJsonForInvoice_(
    (receiptEnriched.billingJson || []).filter(row => !(row && row.skipInvoice)),
    targetPatientIds
  );
  const invoiceContexts = targetBillingRows.map(row => {
    const pid = billingNormalizePatientId_(row && row.patientId);
    const receiptEntry = pid
      ? (receiptEnriched.billingJson || []).find(item => billingNormalizePatientId_(item && item.patientId) === pid) || row
      : row;
    return buildInvoicePdfContextForEntry_(receiptEntry, receiptEnriched, monthCache);
  }).filter(Boolean);
  // 'scheduled' represents bank-flag-driven aggregate entries that skipped standard
  // invoices earlier and are now eligible for aggregate PDFs without additional triggers.
  const aggregateTargets = (receiptEnriched.billingJson || []).filter(
    row => row
      && row.skipInvoice
      && row.bankFlags
      && row.bankFlags.af === true
  );
  const aggregateFiles = aggregateTargets.map(entry => {
    const pid = billingNormalizePatientId_(entry && entry.patientId);
    if (!pid) return null;

    const baseEntry = (receiptEnriched.billingJson || []).find(row => billingNormalizePatientId_(row && row.patientId) === pid)
      || (normalized.billingJson || []).find(row => billingNormalizePatientId_(row && row.patientId) === pid)
      || entry;
    const aggregateUntilMonth = normalizeBillingMonthKeySafe_(baseEntry.aggregateUntilMonth || normalized.aggregateUntilMonth);
    const aggregateSourceMonths = collectAggregateBankFlagMonthsForPatient_(normalized.billingMonth, pid, aggregateUntilMonth, monthCache);
    const aggregateMonths = normalizeAggregateInvoiceMonths_(aggregateSourceMonths, normalized, normalized.billingMonth);
    const uniqueAggregateMonths = Array.from(new Set(aggregateMonths))
      .filter(month => month !== normalized.billingMonth);
    const aggregateTotal = uniqueAggregateMonths.reduce(
      (sum, ym) => sum + resolveBillingAmountForMonthAndPatient_(
        ym,
        pid,
        null,
        monthCache
      ),
      0
    );
    const aggregateRemark = formatAggregateBillingRemark_(uniqueAggregateMonths);
    const aggregateEntry = Object.assign({}, baseEntry, {
      billingMonth: normalized.billingMonth,
      receiptMonths: uniqueAggregateMonths,
      aggregateRemark,
      billingAmount: aggregateTotal,
      transportAmount: 0,
      total: aggregateTotal,
      grandTotal: aggregateTotal,
      aggregateTargetMonths: uniqueAggregateMonths
    });
    const aggregateContext = buildAggregateInvoicePdfContext_(aggregateEntry, uniqueAggregateMonths, receiptEnriched, monthCache);
    const meta = generateAggregateInvoicePdf(aggregateContext, { billingMonth: normalized.billingMonth });
    return Object.assign({}, meta, { patientId: aggregateEntry.patientId, nameKanji: aggregateEntry.nameKanji });
  }).filter(Boolean);
  const matchedIds = new Set(
    targetBillingRows
      .concat(aggregateFiles.map(file => ({ patientId: file && file.patientId })))
      .map(row => String(row && row.patientId ? row.patientId : '').trim())
      .filter(Boolean)
  );
  const missingPatientIds = targetPatientIds.filter(id => !matchedIds.has(id));

  const outputOptions = Object.assign({}, options, { billingMonth: normalized.billingMonth, patientIds: targetPatientIds });
  const pdfs = generateInvoicePdfs(invoiceContexts, outputOptions);
  const shouldExportBank = !outputOptions || outputOptions.skipBankExport !== true;
  const bankOutput = shouldExportBank ? exportBankTransferDataForPrepared_(aggregateApplied) : null;
  return {
    billingMonth: normalized.billingMonth,
    billingJson: receiptEnriched.billingJson,
    receiptStatus: normalized.receiptStatus || '',
    aggregateUntilMonth: normalized.aggregateUntilMonth || '',
    files: pdfs.files.concat(aggregateFiles),
    invoicePatientIds: targetPatientIds,
    missingInvoicePatientIds: missingPatientIds,
    invoiceGenerationMode: targetPatientIds.length ? 'partial' : 'bulk',
    bankOutput,
    preparedAt: normalized.preparedAt || null
  };
}

function generateInvoicesFromCache(billingMonth, options) {
  const month = normalizeBillingMonthInput(billingMonth);
  const loaded = loadPreparedBillingWithSheetFallback_(month.key, { withValidation: true });
  const validation = loaded && loaded.validation ? loaded.validation : null;
  const prepared = normalizePreparedBilling_(loaded && loaded.prepared);
  if (!prepared || !prepared.billingJson || (validation && validation.ok === false)) {
    throw new Error('事前集計が見つかりません。先に「請求データを集計」ボタンを実行してください。');
  }
  return generatePreparedInvoices_(prepared, options);
}

function normalizeAggregateInvoiceMonths_(months, prepared, billingMonth) {
  const aggregateUntilMonth = prepared && prepared.aggregateUntilMonth ? prepared.aggregateUntilMonth : '';
  const source = []
    .concat(months || [])
    .concat(aggregateUntilMonth ? [aggregateUntilMonth] : []);
  const normalized = normalizePastBillingMonths_(source, billingMonth || (prepared && prepared.billingMonth));
  if (typeof normalizeAggregateMonthsForInvoice_ === 'function') {
    return normalizeAggregateMonthsForInvoice_(normalized, billingMonth || (prepared && prepared.billingMonth));
  }
  return normalized;
}

function finalizeAggregateInvoiceState_(prepared, billingMonth, patientId) {
  const normalizedPrepared = normalizePreparedBilling_(prepared);
  const billingMonthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = billingNormalizePatientId_(patientId);

  if (!normalizedPrepared || !billingMonthKey || !pid || !Array.isArray(normalizedPrepared.billingJson)) return null;

  const index = normalizedPrepared.billingJson.findIndex(row => billingNormalizePatientId_(row && row.patientId) === pid);
  if (index < 0) return null;

  const targetEntry = normalizedPrepared.billingJson[index] || {};
  const existingAggregateUntil = normalizeBillingMonthKeySafe_(
    targetEntry.aggregateUntilMonth || normalizedPrepared.aggregateUntilMonth
  );
  const existingNum = Number(existingAggregateUntil) || 0;
  const targetNum = Number(billingMonthKey) || 0;
  const resolvedAggregateUntil = targetNum > existingNum ? billingMonthKey : existingAggregateUntil;
  const currentReceiptStatus = normalizeReceiptStatus_(targetEntry.receiptStatus || normalizedPrepared.receiptStatus);
  const needsReceiptStatusUpdate = currentReceiptStatus !== 'AGGREGATE';
  const needsAggregateUpdate = resolvedAggregateUntil && resolvedAggregateUntil !== existingAggregateUntil;

  if (!needsReceiptStatusUpdate && !needsAggregateUpdate) return null;

  const updatedEntry = Object.assign({}, targetEntry, {
    receiptStatus: 'AGGREGATE',
    aggregateUntilMonth: resolvedAggregateUntil
  });

  const updatedBillingJson = normalizedPrepared.billingJson.slice();
  updatedBillingJson[index] = updatedEntry;

  return Object.assign({}, normalizedPrepared, { billingJson: updatedBillingJson });
}

function generateAggregatedInvoice(billingMonth, options) {
  const month = normalizeBillingMonthInput(billingMonth);
  const opts = options || {};
  const patientId = billingNormalizePatientId_(opts.patientId);
  if (!patientId) {
    throw new Error('患者IDを指定してください');
  }

  const loaded = loadPreparedBillingWithSheetFallback_(month.key, { withValidation: true, restoreCache: true });
  const validation = loaded && loaded.validation ? loaded.validation : null;
  const prepared = normalizePreparedBilling_(loaded && loaded.prepared);
  if (!prepared || !prepared.billingJson || (validation && validation.ok === false)) {
    throw new Error('事前集計が見つかりません。先に「請求データを集計」ボタンを実行してください。');
  }

  const entry = (prepared.billingJson || []).find(row => billingNormalizePatientId_(row && row.patientId) === patientId);
  if (!entry) {
    throw new Error('対象患者の請求データが見つかりません');
  }

  const mergedMonths = []
    .concat(opts.aggregateMonths || [])
    .concat(entry.aggregateMonths || [])
    .concat(entry.receiptMonths || []);
  const aggregateMonths = normalizeAggregateInvoiceMonths_(mergedMonths, prepared, month.key);

  const monthCache = createBillingMonthCache_();
  const receiptPrepared = attachPreviousReceiptAmounts_(prepared, monthCache);
  const receiptEntry = (receiptPrepared && receiptPrepared.billingJson || [])
    .find(row => billingNormalizePatientId_(row && row.patientId) === patientId) || entry;
  const aggregateContext = buildAggregateInvoicePdfContext_(receiptEntry, aggregateMonths, receiptPrepared || prepared, monthCache);
  const file = generateAggregateInvoicePdf(aggregateContext, { billingMonth: month.key });
  return { billingMonth: month.key, patientId, aggregateMonths, file };
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

function generatePreparedInvoicesForMonth(billingMonth, options) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) {
    throw new Error('PDF対象月が指定されていません。');
  }

  const opts = options || {};
  if (opts.applyEdits === true) {
    applyBillingEdits(monthKey, opts);
  }

  const loaded = loadPreparedBilling_(monthKey, { withValidation: true });
  const validation = loaded && loaded.validation ? loaded.validation : null;
  const prepared = normalizePreparedBilling_(loaded && loaded.prepared);
  if (!prepared || !prepared.billingJson || (validation && validation.ok === false)) {
    throw new Error('事前集計が見つかりません。先に「請求データを集計」ボタンを実行してください。');
  }

  return generatePreparedInvoices_(prepared, opts);
}

function updateBillingReceiptStatus(billingMonth, options) {
  const month = normalizeBillingMonthInput(billingMonth);
  const status = normalizeReceiptStatus_(options && options.receiptStatus);
  const aggregateUntil = status === 'AGGREGATE'
    ? normalizeBillingMonthKeySafe_(options && (options.aggregateUntil || options.aggregateUntilMonth))
    : '';

  const existing = loadPreparedBillingWithSheetFallback_(month.key, { allowInvalid: true, restoreCache: false });
  let prepared = existing && existing.prepared !== undefined ? existing.prepared : existing;
  if (!prepared) {
    prepared = buildPreparedBillingPayload_(month);
  }

  const merged = mergeReceiptSettingsIntoPrepared_(prepared, status, aggregateUntil);
  savePreparedBilling_(merged);
  savePreparedBillingToSheet_(month.key, merged);
  return toClientBillingPayload_(merged);
}

  // Deprecated (legacy bank transfer export). New specifications do not rely on bank CSV/JSON outputs.
  function generateBankTransferData(billingMonth, options) {
    const prepared = buildPreparedBillingPayload_(billingMonth);
    savePreparedBilling_(prepared);
    savePreparedBillingToSheet_(billingMonth, prepared);
    return exportBankTransferDataForPrepared_(prepared, options || {});
  }

  // Deprecated (legacy bank transfer export). New specifications do not rely on bank CSV/JSON outputs.
  function generateBankTransferDataFromCache(billingMonth, options) {
    const opts = options || {};
    const monthInput = billingMonth || opts.billingMonth || (opts.prepared && (opts.prepared.billingMonth || opts.prepared.month));
    const resolvedMonthKey = normalizeBillingMonthKeySafe_(monthInput);
    const month = normalizeBillingMonthInput
      ? normalizeBillingMonthInput(resolvedMonthKey || monthInput)
      : null;
    const loaded = month && month.key
      ? loadPreparedBillingWithSheetFallback_(month.key, { withValidation: true, restoreCache: true })
      : null;
    const validation = loaded && loaded.validation
      ? loaded.validation
      : (loaded && loaded.ok !== undefined ? loaded : null);
    const normalizedPrepared = normalizePreparedBilling_(loaded && loaded.prepared);

    try {
      billingLogger_.log('[bankExport] generateBankTransferDataFromCache summary=' + JSON.stringify({
        requestedMonth: monthInput || null,
        resolvedMonthKey: month ? month.key : null,
        preparedBillingMonth: normalizedPrepared && normalizedPrepared.billingMonth ? normalizedPrepared.billingMonth : null,
        preparedBillingJsonLength: normalizedPrepared && Array.isArray(normalizedPrepared.billingJson)
          ? normalizedPrepared.billingJson.length
          : null,
        validationBillingMonth: validation && validation.billingMonth ? validation.billingMonth : null,
        validationReason: validation && validation.reason ? validation.reason : null,
        validationOk: validation && Object.prototype.hasOwnProperty.call(validation, 'ok') ? validation.ok : null
      }));
    } catch (err) {
      // ignore logging errors in non-GAS environments
    }

    if (!month) {
      throw new Error('銀行データを出力できません。請求月が指定されていません。先に請求データを集計してください。');
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
  const missingReasons = ['payload missing', 'billingMonth missing'];
  const corruptReasons = ['billingJson empty', 'billingJson missing'];
  try {
    billingLogger_.log('[bankExport] resolveBankExportErrorMessage_ reason=' + reason);
  } catch (err) {
    // ignore logging errors in non-GAS environments
  }
  if (ledgerReasons.indexOf(reason) >= 0) {
    return '銀行データを生成できません。繰越金データを確認してください。';
  }
  if (missingReasons.indexOf(reason) >= 0 || !reason) {
    return '銀行データを出力できません。請求データが見つかりません。先に請求データを集計してください。';
  }
  if (corruptReasons.indexOf(reason) >= 0) {
    return '銀行データを生成できません。請求データが破損しています。再集計してください。';
  }
  if (reason === 'billingMonth mismatch') {
    return '銀行データを生成できません。請求月が一致する請求データを集計してください。';
  }
  return '銀行データを生成できません。請求データは存在しますが、検証に失敗しました。';
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

function billingApplyUnpaidFromBankSheetFromMenu() {
  const month = promptBillingMonthInput_();
  const result = applyBankWithdrawalUnpaidEntries(month.key);
  const lines = [
    '未回収履歴へ反映しました',
    'チェック済み行: ' + result.checkedRows,
    '追加件数: ' + (result.added || 0)
  ];
  if (result.skipped) {
    lines.push('スキップ: ' + result.skipped);
  }
  SpreadsheetApp.getUi().alert(lines.join('\n'));
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
  const workbook = billingSs();
  if (!workbook) {
    return { billingMonth, exists: false };
  }
  const sheet = workbook.getSheetByName('請求履歴');
  if (!sheet) {
    return { billingMonth, exists: false };
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { billingMonth, exists: true, total: 0, statuses: {}, paidTotal: 0, unpaidTotal: 0 };
  }
  const colCount = Math.max(sheet.getLastColumn(), 11);
  const headers = sheet.getRange(1, 1, 1, colCount).getDisplayValues()[0];
  const columns = typeof resolveBillingHistoryColumnsFromHeaders_ === 'function'
    ? resolveBillingHistoryColumnsFromHeaders_(headers)
    : {
      billingMonth: 1,
      paidAmount: 7,
      unpaidAmount: 8,
      bankStatus: 9
    };
  const billingMonthIdx = (columns.billingMonth || 1) - 1;
  const paidIdx = (columns.paidAmount || 7) - 1;
  const unpaidIdx = (columns.unpaidAmount || 8) - 1;
  const bankStatusIdx = (columns.bankStatus || 9) - 1;

  const values = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
  const rows = values.filter(row => row[billingMonthIdx] === billingMonth);
  const statuses = {};
  let paidTotal = 0;
  let unpaidTotal = 0;
  rows.forEach(row => {
    const status = row[bankStatusIdx] || '';
    statuses[status] = (statuses[status] || 0) + 1;
    paidTotal += Number(row[paidIdx]) || 0;
    unpaidTotal += Number(row[unpaidIdx]) || 0;
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
  billingMenu.addItem('銀行引落の未回収を履歴へ反映', 'billingApplyUnpaidFromBankSheetFromMenu');
  billingMenu.addToUi();

  const attendanceMenu = ui.createMenu('勤怠管理');
  attendanceMenu.addItem('勤怠データを今すぐ同期', 'runVisitAttendanceSyncJobFromMenu');
  attendanceMenu.addItem('日次同期トリガーを確認', 'ensureVisitAttendanceSyncTriggerFromMenu');
  attendanceMenu.addToUi();
}
