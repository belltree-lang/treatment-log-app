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
  const normalized = normalizeBillingEntryFromEntries_(source);
  const insuranceEntry = resolveBillingEntryByType_(normalized, 'insurance') || {};
  const selfPayEntries = resolveBillingEntries_(normalized).filter(item => (
    normalizeBillingEntryTypeValue_(item && (item.type || item.entryType)) === 'self_pay'
  ));
  const selfPayEntryWithManual = selfPayEntries.find(item => item && item.manualOverride
    && item.manualOverride.amount !== '' && item.manualOverride.amount !== null
    && item.manualOverride.amount !== undefined);
  const manualUnitPrice = Object.prototype.hasOwnProperty.call(source, 'manualUnitPrice')
    ? source.manualUnitPrice
    : insuranceEntry.manualUnitPrice;
  const manualTransportAmount = Object.prototype.hasOwnProperty.call(source, 'manualTransportAmount')
    ? source.manualTransportAmount
    : (Object.prototype.hasOwnProperty.call(insuranceEntry, 'manualTransportAmount')
      ? insuranceEntry.manualTransportAmount
      : insuranceEntry.transportAmount);
  const manualSelfPayAmount = Object.prototype.hasOwnProperty.call(source, 'manualSelfPayAmount')
    ? source.manualSelfPayAmount
    : (selfPayEntryWithManual && selfPayEntryWithManual.manualOverride
      && Object.prototype.hasOwnProperty.call(selfPayEntryWithManual.manualOverride, 'amount')
      ? selfPayEntryWithManual.manualOverride.amount
      : undefined);
  const selfPayItems = selfPayEntries.reduce((list, entry) => {
    const items = Array.isArray(entry && entry.items)
      ? entry.items
      : (Array.isArray(entry && entry.selfPayItems) ? entry.selfPayItems : []);
    return list.concat(items);
  }, []);
  const amountCalc = calculateBillingAmounts_({
    visitCount: insuranceEntry.visitCount,
    insuranceType: insuranceEntry.insuranceType,
    burdenRate: insuranceEntry.burdenRate,
    manualUnitPrice: manualUnitPrice != null ? manualUnitPrice : insuranceEntry.unitPrice,
    manualTransportAmount,
    unitPrice: insuranceEntry.unitPrice,
    medicalAssistance: insuranceEntry.medicalAssistance,
    carryOverAmount: insuranceEntry.carryOverAmount,
    selfPayItems,
    manualSelfPayAmount
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
const BANK_WITHDRAWAL_ONLINE_HEADER = 'オンライン';
const BILLING_DEBUG_PID = '';

if (typeof globalThis !== 'undefined') {
  globalThis.BILLING_CACHE_CHUNK_MARKER = BILLING_CACHE_CHUNK_MARKER;
  globalThis.BILLING_CACHE_MAX_ENTRY_LENGTH = BILLING_CACHE_MAX_ENTRY_LENGTH;
  globalThis.BILLING_CACHE_CHUNK_SIZE = BILLING_CACHE_CHUNK_SIZE;
}

const BILLING_MONTH_KEY_CACHE_ = {};
const RECEIPT_TARGET_MONTHS_BY_BANK_FLAGS_CACHE_ = {};
const BILLING_CACHE_PAYLOAD_MEMO_ = {};

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

function normalizeBillingMonthKeyText_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  if (/^\d{6}$/.test(raw)) {
    return raw;
  }
  const normalizedDigits = raw.replace(/\D/g, '');
  if (normalizedDigits.length === 6) {
    return normalizedDigits;
  }
  const match = raw.match(/^(\d{4})\s*[\/-]?\s*(\d{1,2})$/);
  if (match) {
    const yearNum = Number(match[1]);
    const monthNum = Number(match[2]);
    if (Number.isFinite(yearNum) && Number.isFinite(monthNum) && monthNum >= 1 && monthNum <= 12) {
      return String(yearNum).padStart(4, '0') + String(monthNum).padStart(2, '0');
    }
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
    let normalized = '';
    if (candidate instanceof Date && !isNaN(candidate.getTime())) {
      const year = String(candidate.getFullYear()).padStart(4, '0');
      const month = String(candidate.getMonth() + 1).padStart(2, '0');
      normalized = year + month;
    } else if (candidate && typeof candidate === 'object') {
      if (candidate.key) {
        normalized = normalizeBillingMonthKeyText_(candidate.key);
      } else if (candidate.billingMonth) {
        normalized = normalizeBillingMonthKeyText_(candidate.billingMonth);
      } else if (candidate.ym) {
        normalized = normalizeBillingMonthKeyText_(candidate.ym);
      } else if (candidate.month && candidate.month.key) {
        normalized = normalizeBillingMonthKeyText_(candidate.month.key);
      } else if (candidate.month && candidate.month.ym) {
        normalized = normalizeBillingMonthKeyText_(candidate.month.ym);
      } else if (candidate.year && candidate.month) {
        const yearNum = Number(candidate.year);
        const monthNum = Number(candidate.month);
        if (Number.isFinite(yearNum) && Number.isFinite(monthNum)) {
          normalized = String(yearNum).padStart(4, '0') + String(monthNum).padStart(2, '0');
        }
      }
    } else if (typeof candidate === 'string' || typeof candidate === 'number') {
      normalized = normalizeBillingMonthKeyText_(candidate);
    }
    if (normalized) {
      if (cacheKey) {
        BILLING_MONTH_KEY_CACHE_[cacheKey] = normalized;
      }
      return normalized;
    }
    const fallback = String(candidate || '').trim();
    if (fallback) {
      if (cacheKey) {
        BILLING_MONTH_KEY_CACHE_[cacheKey] = fallback;
      }
      return fallback;
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
  const existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];
  const existingIndexes = [];
  existingRows.forEach((row, idx) => {
    if (String(row[0] || '').trim() === monthKey) {
      existingIndexes.push(idx);
    }
  });

  if (!metaPayload || typeof metaPayload !== 'object') {
    let cleared = 0;
    if (existingIndexes.length) {
      existingIndexes.forEach(idx => {
        sheet.getRange(idx + 2, 1, 1, 3).setValues([['', '', '']]);
        cleared += 1;
      });
    }
    return { billingMonth: monthKey, inserted: 0, updated: 0, appended: 0, cleared };
  }

  const payloadJson = JSON.stringify(metaPayload);
  const chunkSize = 40000;
  const chunks = [];
  for (let i = 0; i < payloadJson.length; i += chunkSize) {
    chunks.push(payloadJson.slice(i, i + chunkSize));
  }

  if (!chunks.length) return { billingMonth: monthKey, inserted: 0 };

  const rows = chunks.map((chunk, idx) => [monthKey, idx + 1, chunk]);
  const updateCount = Math.min(existingIndexes.length, rows.length);
  for (let i = 0; i < updateCount; i++) {
    const rowIndex = existingIndexes[i] + 2;
    sheet.getRange(rowIndex, 1, 1, 3).setValues([rows[i]]);
  }

  const appendCount = rows.length - updateCount;
  if (appendCount > 0) {
    const appendStart = sheet.getLastRow() + 1;
    sheet.getRange(appendStart, 1, appendCount, 3).setValues(rows.slice(updateCount));
  }

  let cleared = 0;
  if (existingIndexes.length > rows.length) {
    for (let i = rows.length; i < existingIndexes.length; i++) {
      const rowIndex = existingIndexes[i] + 2;
      sheet.getRange(rowIndex, 1, 1, 3).setValues([['', '', '']]);
      cleared += 1;
    }
  }

  return { billingMonth: monthKey, inserted: rows.length, updated: updateCount, appended: appendCount, cleared };
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
    bankWithdrawalAmountsByMonth: {},
    preparedByMonthLoadedAll: false,
    bankWithdrawalAmountsLoadedAll: false
  };
}

function getPreparedPayloadRuntimeCache_(cache) {
  if (!cache || !cache.preparedEntriesByMonth) return null;
  return cache.preparedEntriesByMonth;
}

function storePreparedPayloadInRuntimeCache_(prepared, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(prepared && prepared.billingMonth);
  const store = getPreparedPayloadRuntimeCache_(cache);
  if (!monthKey || !store) return;
  store[monthKey] = normalizePreparedBilling_(prepared);
}

function getPreparedPayloadForMonthCached_(billingMonth, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const store = getPreparedPayloadRuntimeCache_(cache);
  if (!monthKey || !store || !Object.prototype.hasOwnProperty.call(store, monthKey)) return null;
  return store[monthKey] || null;
}

function getPreparedBillingForMonthCached_(billingMonth, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) return null;
  const store = cache && cache.preparedByMonth ? cache.preparedByMonth : null;
  if (!store) return null;
  // NOTE: This function never loads from sheets. Callers must preload cache
  // via loadPreparedBillingSummariesIntoCache_.
  if (Object.prototype.hasOwnProperty.call(store, monthKey)) {
    return store[monthKey];
  }
  return null;
}

function getPreparedBillingEntryForMonthCached_(billingMonth, patientId, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = billingNormalizePatientId_(patientId);
  if (!monthKey || !pid) return null;
  const runtimePayload = getPreparedPayloadForMonthCached_(monthKey, cache);
  if (runtimePayload && Array.isArray(runtimePayload.billingJson)) {
    const runtimeEntry = runtimePayload.billingJson
      .find(entry => billingNormalizePatientId_(entry && entry.patientId) === pid);
    if (runtimeEntry) {
      return pickPreparedBillingEntrySummary_(runtimeEntry);
    }
  }
  // NOTE: Cache must be preloaded (loadPreparedBillingSummariesIntoCache_).
  const summary = getPreparedBillingForMonthCached_(monthKey, cache);
  const totals = summary && summary.totalsByPatient ? summary.totalsByPatient : null;
  if (!totals || !Object.prototype.hasOwnProperty.call(totals, pid)) return null;
  return totals[pid] || null;
}

function collectPreparedPayloadMonthsForPdf_(prepared, cache) {
  const normalized = normalizePreparedBilling_(prepared);
  const monthKey = normalizeBillingMonthKeySafe_(normalized && normalized.billingMonth);
  if (!monthKey) return [];

  const monthSet = new Set([monthKey]);
  if (!normalized || !Array.isArray(normalized.billingJson)) return Array.from(monthSet);

  normalized.billingJson.forEach(entry => {
    const pid = billingNormalizePatientId_(entry && entry.patientId);
    if (!pid) return;

    const receiptMonths = resolveReceiptTargetMonths(pid, monthKey, cache);
    receiptMonths.forEach(targetMonth => {
      const normalizedMonth = normalizeBillingMonthKeySafe_(targetMonth);
      if (normalizedMonth) monthSet.add(normalizedMonth);
    });

    const decision = resolveInvoiceGenerationMode(pid, monthKey, cache);
    const decisionMonths = decision && Array.isArray(decision.aggregateMonths) ? decision.aggregateMonths : [];
    decisionMonths.forEach(targetMonth => {
      const normalizedMonth = normalizeBillingMonthKeySafe_(targetMonth);
      if (normalizedMonth) monthSet.add(normalizedMonth);
    });

    const entryMonths = []
      .concat(entry.aggregateMonths || [])
      .concat(entry.receiptMonths || [])
      .concat(entry.aggregateTargetMonths || []);
    entryMonths.forEach(targetMonth => {
      const normalizedMonth = normalizeBillingMonthKeySafe_(targetMonth);
      if (normalizedMonth) monthSet.add(normalizedMonth);
    });

    const aggregateUntilMonth = normalizeBillingMonthKeySafe_(
      entry.aggregateUntilMonth || normalized.aggregateUntilMonth
    );
    const aggregateSourceMonths = collectAggregateBankFlagMonthsForPatient_(
      monthKey,
      pid,
      aggregateUntilMonth,
      cache
    );
    aggregateSourceMonths.forEach(targetMonth => {
      const normalizedMonth = normalizeBillingMonthKeySafe_(targetMonth);
      if (normalizedMonth) monthSet.add(normalizedMonth);
    });
  });

  return Array.from(monthSet).filter(Boolean).sort();
}

function preloadPreparedPayloadsForPdfGeneration_(prepared, cache) {
  const normalized = normalizePreparedBilling_(prepared);
  const monthKey = normalizeBillingMonthKeySafe_(normalized && normalized.billingMonth);
  const store = getPreparedPayloadRuntimeCache_(cache);
  if (!monthKey || !store) return;

  if (!cache.preparedByMonthLoadedAll) {
    loadPreparedBillingSummariesIntoCache_(cache);
  }

  storePreparedPayloadInRuntimeCache_(normalized, cache);

  const monthsToLoad = collectPreparedPayloadMonthsForPdf_(normalized, cache)
    .filter(targetMonth => targetMonth !== monthKey);

  monthsToLoad.forEach(targetMonth => {
    if (Object.prototype.hasOwnProperty.call(store, targetMonth)) return;
    const payload = loadPreparedBillingFromSheet_(targetMonth);
    const validation = validatePreparedBillingPayload_(payload, targetMonth);
    if (payload && validation && validation.ok) {
      store[targetMonth] = normalizePreparedBilling_(Object.assign({}, payload, { billingMonth: targetMonth }));
    } else {
      store[targetMonth] = null;
    }
  });
}

function loadPreparedBillingSummariesIntoCache_(cache) {
  const store = cache && cache.preparedByMonth ? cache.preparedByMonth : null;
  if (!store || cache.preparedByMonthLoadedAll) return;
  const months = getPreparedBillingMonths();
  months.forEach(monthKey => {
    if (Object.prototype.hasOwnProperty.call(store, monthKey)) return;
    const summary = loadPreparedBillingSummaryFromSheet_(monthKey);
    const reduced = reducePreparedBillingSummary_(summary);
    store[monthKey] = reduced || null;
  });
  cache.preparedByMonthLoadedAll = true;
}

function loadBankWithdrawalAmountsIntoCache_(cache, prepared) {
  const store = cache && cache.bankWithdrawalAmountsByMonth ? cache.bankWithdrawalAmountsByMonth : null;
  if (!store || cache.bankWithdrawalAmountsLoadedAll) return;
  const months = cache && cache.preparedByMonth ? Object.keys(cache.preparedByMonth) : [];
  months.forEach(monthKey => {
    if (!monthKey || Object.prototype.hasOwnProperty.call(store, monthKey)) return;
    store[monthKey] = collectBankWithdrawalAmountsByPatient_(monthKey, prepared) || {};
  });
  cache.bankWithdrawalAmountsLoadedAll = true;
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
    billingItems: row.billingItems,
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
  const onlineCol = resolveBillingColumn_(headers, [BANK_WITHDRAWAL_ONLINE_HEADER], BANK_WITHDRAWAL_ONLINE_HEADER, {});
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

function resolveInvoiceGenerationMode(patientId, billingMonth, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(patientId)
    : String(patientId || '').trim();

  if (!monthKey || !pid) {
    return { mode: 'standard', aggregateMonths: [] };
  }

  const currentFlags = getBankWithdrawalStatusByPatient_(monthKey, pid, cache);
  if (!(currentFlags && currentFlags.af)) {
    if (typeof shouldLogReceiptDebug_ === 'function' && shouldLogReceiptDebug_(pid)) {
      Logger.log('[receipt-debug][resolveInvoiceGenerationMode] ' + JSON.stringify({
        patientId: pid,
        billingMonth: monthKey,
        currentFlags,
        priorMonthsChecked: [],
        aeMonths: [],
        resolvedMode: 'standard'
      }));
    }
    return { mode: 'standard', aggregateMonths: [] };
  }

  const aeMonths = [];
  const priorMonthsChecked = [];
  let cursor = resolvePreviousBillingMonthKey_(monthKey);
  let guard = 0;

  while (cursor && guard < 48) {
    const cursorKey = normalizeBillingMonthKeySafe_(cursor);
    if (!cursorKey) break;
    priorMonthsChecked.push(cursorKey);
    const flags = getBankWithdrawalStatusByPatient_(cursorKey, pid, cache);
    if (flags && flags.ae) {
      aeMonths.unshift(cursorKey);
      cursor = resolvePreviousBillingMonthKey_(cursorKey);
      guard += 1;
      continue;
    }
    break;
  }

  const resolvedMode = aeMonths.length ? 'aggregate' : 'standard';
  const aggregateMonths = resolvedMode === 'aggregate'
    ? normalizePastBillingMonths_(aeMonths.concat(monthKey), monthKey)
    : [];

  if (typeof shouldLogReceiptDebug_ === 'function' && shouldLogReceiptDebug_(pid)) {
    Logger.log('[receipt-debug][resolveInvoiceGenerationMode] ' + JSON.stringify({
      patientId: pid,
      billingMonth: monthKey,
      currentFlags,
      priorMonthsChecked,
      aeMonths,
      resolvedMode
    }));
  }

  return { mode: resolvedMode, aggregateMonths };
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

function attachPreviousReceiptAmounts_(prepared, cache, options) {
  const monthKey = prepared && prepared.billingMonth;
  if (!monthKey || !Array.isArray(prepared && prepared.billingJson)) return prepared;

  const previousMonthKey = resolvePreviousBillingMonthKey_(monthKey);
  if (!previousMonthKey) return prepared;

  const normalizePid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_
    : value => String(value || '').trim();
  const opts = options || {};
  const targetIds = Array.isArray(opts.targetPatientIds) ? opts.targetPatientIds : [];
  const normalizedTargetIds = targetIds
    .map(id => normalizePid(id))
    .filter(Boolean);
  const targetIdSet = normalizedTargetIds.length ? new Set(normalizedTargetIds) : null;
  const monthCache = cache || {
    preparedByMonth: {},
    bankWithdrawalUnpaidByMonth: {},
    bankWithdrawalAmountsByMonth: {}
  };

  const enrichedJson = prepared.billingJson.map(entry => {
    const pid = normalizePid(entry && entry.patientId);
    if (targetIdSet && !targetIdSet.has(pid)) {
      return entry;
    }
    const hasPreviousPreparedEntry = !!getPreparedBillingEntryForMonthCached_(previousMonthKey, pid, monthCache);
    const receiptTargetMonths = resolveReceiptTargetMonths(pid, monthKey, cache);
    const hasPreviousReceiptSheet = hasPreviousPreparedEntry;
    let nextEntry = entry;

    if (!receiptTargetMonths.length) {
      logReceiptDebug_(pid, {
        step: 'attachPreviousReceiptAmounts_',
        billingMonth: monthKey,
        patientId: pid,
        receiptTargetMonths,
        receiptMonths: []
      });
    } else if (!hasPreviousReceiptSheet && receiptTargetMonths[0] === previousMonthKey) {
      nextEntry = Object.assign({}, entry, {
        hasPreviousReceiptSheet: false,
        receiptRemark: '',
        receiptMonthBreakdown: []
      });
    } else {
      const receiptBreakdown = buildReceiptMonthBreakdownForEntry_(pid, receiptTargetMonths, prepared, monthCache);

      logReceiptDebug_(pid, {
        step: 'attachPreviousReceiptAmounts_',
        billingMonth: monthKey,
        patientId: pid,
        receiptTargetMonths,
        receiptMonths: receiptTargetMonths
      });
      nextEntry = Object.assign({}, entry, {
        hasPreviousReceiptSheet,
        receiptMonths: receiptTargetMonths,
        receiptRemark: entry && entry.receiptRemark,
        receiptMonthBreakdown: receiptBreakdown
      });
    }

    if (nextEntry && nextEntry.previousReceiptAmount == null) {
      if (nextEntry && Array.isArray(nextEntry.receiptMonthBreakdown) && nextEntry.receiptMonthBreakdown.length) {
        const hasBreakdownAmount = nextEntry.receiptMonthBreakdown.some(item => item && item.amount != null && item.amount !== '');
        if (hasBreakdownAmount) {
          const breakdownAmount = nextEntry.receiptMonthBreakdown.reduce((sum, item) => {
            const normalized = normalizeMoneyNumber_(item && item.amount);
            return normalized != null ? sum + normalized : sum;
          }, 0);
          if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
            billingLogger_.log('[billing] buildInvoicePdfContextForEntry_ filled previousReceiptAmount from receiptMonthBreakdown: ' + JSON.stringify({
              patientId: pid,
              billingMonth: monthKey,
              breakdownAmount
            }));
          }
          return Object.assign({}, nextEntry, { previousReceiptAmount: breakdownAmount });
        }
      }
      const previousEntry = getPreparedBillingEntryForMonthCached_(previousMonthKey, pid, monthCache);
      if (!previousEntry) return nextEntry;
      const previousAmount = previousEntry.grandTotal != null && previousEntry.grandTotal !== ''
        ? normalizeMoneyNumber_(previousEntry.grandTotal)
        : normalizeMoneyNumber_(previousEntry.billingAmount);
      if (previousAmount == null) return nextEntry;
      if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
        billingLogger_.log('[billing] buildInvoicePdfContextForEntry_ filled previousReceiptAmount from previous month: ' + JSON.stringify({
          patientId: pid,
          billingMonth: monthKey,
          previousMonthKey,
          previousAmount
        }));
      }
      return Object.assign({}, nextEntry, { previousReceiptAmount: previousAmount });
    }

    return nextEntry;
  });

  return Object.assign({}, prepared, {
    billingJson: enrichedJson,
    hasPreviousReceiptSheet: !!getPreparedBillingForMonthCached_(previousMonthKey, monthCache)
  });
}

function buildReceiptSummaryMap_(prepared, cache, options) {
  const monthKey = normalizeBillingMonthKeySafe_(prepared && prepared.billingMonth);
  if (!monthKey || !Array.isArray(prepared && prepared.billingJson)) return {};

  const normalizePid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_
    : value => String(value || '').trim();
  const opts = options || {};
  const targetIds = Array.isArray(opts.targetPatientIds) ? opts.targetPatientIds : [];
  const normalizedTargetIds = targetIds
    .map(id => normalizePid(id))
    .filter(Boolean);
  const targetIdSet = normalizedTargetIds.length ? new Set(normalizedTargetIds) : null;
  const map = {};

  prepared.billingJson.forEach(entry => {
    const pid = normalizePid(entry && entry.patientId);
    if (!pid) return;
    if (targetIdSet && !targetIdSet.has(pid)) return;
    const decision = resolveInvoiceGenerationMode(pid, monthKey, cache);
    const decisionMonths = decision && Array.isArray(decision.aggregateMonths) ? decision.aggregateMonths : [];
    const aggregateMonths = normalizePastBillingMonths_(decisionMonths, monthKey);
    const isAggregateInvoice = !!(decision && decision.mode === 'aggregate' && aggregateMonths.length > 1);
    map[pid] = {
      decision,
      aggregateMonths,
      isAggregateInvoice
    };
  });

  return map;
}

function resolveReceiptSummaryForPatient_(patientId, billingMonth, cache, summaryMap) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(patientId)
    : String(patientId || '').trim();
  if (!pid || !monthKey) {
    return { decision: { mode: 'standard', aggregateMonths: [] }, aggregateMonths: [], isAggregateInvoice: false };
  }
  if (summaryMap && Object.prototype.hasOwnProperty.call(summaryMap, pid)) {
    return summaryMap[pid];
  }
  const decision = resolveInvoiceGenerationMode(pid, monthKey, cache);
  const decisionMonths = decision && Array.isArray(decision.aggregateMonths) ? decision.aggregateMonths : [];
  const aggregateMonths = normalizePastBillingMonths_(decisionMonths, monthKey);
  const isAggregateInvoice = !!(decision && decision.mode === 'aggregate' && aggregateMonths.length > 1);
  return {
    decision,
    aggregateMonths,
    isAggregateInvoice
  };
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
    return {};
  }
  return store[monthKey] || {};
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
      const previousEntry = getPreparedBillingEntryForMonthCached_(monthKey, pid, store);
      if (previousEntry) {
        const amount = previousEntry.grandTotal != null && previousEntry.grandTotal !== ''
          ? normalizeMoneyNumber_(previousEntry.grandTotal)
          : normalizeMoneyNumber_(previousEntry.billingAmount);
        if (Number.isFinite(amount)) {
          breakdown.push({ month: monthKey, amount });
        }
      }
    }
  });

  return breakdown;
}

function collectBankWithdrawalAmountsByPatient_(billingMonth, prepared) {
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
  const entries = resolveBillingEntries_(entry);
  if (entries.length) {
    return entries.reduce((sum, billingEntry) => sum + resolveBillingEntryTotalAmount_(billingEntry), 0);
  }
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

function resolveAggregateEntryTotalsForMonths_(aggregateMonths, patientId, fallbackEntry, cache) {
  const pid = billingNormalizePatientId_(patientId);
  const months = normalizePastBillingMonths_(
    Array.isArray(aggregateMonths) ? aggregateMonths : [],
    fallbackEntry && fallbackEntry.billingMonth
  );
  if (!pid || !months.length) {
    return {
      visitCount: 0,
      treatmentAmount: 0,
      transportAmount: 0,
      billingAmount: 0,
      total: 0,
      grandTotal: 0
    };
  }

  const fallbackMonthKey = normalizeBillingMonthKeySafe_(fallbackEntry && fallbackEntry.billingMonth);
  const normalizeVisits = value => (typeof billingNormalizeVisitCount_ === 'function'
    ? billingNormalizeVisitCount_(value)
    : (Number(value) || 0));

  return months.reduce((sum, monthKey) => {
    const normalizedMonth = normalizeBillingMonthKeySafe_(monthKey);
    if (!normalizedMonth) return sum;
    const entry = (fallbackEntry && fallbackMonthKey && normalizedMonth === fallbackMonthKey)
      ? pickPreparedBillingEntrySummary_(fallbackEntry)
      : getPreparedBillingEntryForMonthCached_(normalizedMonth, pid, cache);
    if (!entry) return sum;

    const fallbackAmount = resolveBillingAmountForEntry_(entry);
    const totalValue = entry.total != null && entry.total !== ''
      ? normalizeMoneyNumber_(entry.total)
      : fallbackAmount;
    const grandTotalValue = entry.grandTotal != null && entry.grandTotal !== ''
      ? normalizeMoneyNumber_(entry.grandTotal)
      : fallbackAmount;

    sum.visitCount += normalizeVisits(entry.visitCount);
    sum.treatmentAmount += normalizeMoneyNumber_(entry.treatmentAmount);
    sum.transportAmount += normalizeMoneyNumber_(entry.transportAmount);
    sum.billingAmount += normalizeMoneyNumber_(entry.billingAmount);
    sum.total += totalValue;
    sum.grandTotal += grandTotalValue;
    return sum;
  }, {
    visitCount: 0,
    treatmentAmount: 0,
    transportAmount: 0,
    billingAmount: 0,
    total: 0,
    grandTotal: 0
  });
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

    const preparedEntry = getPreparedBillingEntryForMonthCached_(cursorKey, pid, cache);
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

function resolveReceiptTargetMonthsByBankFlags_(patientId, billingMonth, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(patientId)
    : String(patientId || '').trim();
  if (!monthKey || !pid) return [];

  const cacheKey = pid + ':' + monthKey;
  if (Object.prototype.hasOwnProperty.call(RECEIPT_TARGET_MONTHS_BY_BANK_FLAGS_CACHE_, cacheKey)) {
    return RECEIPT_TARGET_MONTHS_BY_BANK_FLAGS_CACHE_[cacheKey].slice();
  }

  const result = (function resolveReceiptTargetMonthsByBankFlagsWithCache_() {
    const currentFlags = getBankWithdrawalStatusByPatient_(monthKey, pid, cache);
    if (currentFlags && currentFlags.ae) return [];
    if (currentFlags && currentFlags.af) return [];

    const previousMonthKey = resolvePreviousBillingMonthKey_(monthKey);
    if (!previousMonthKey) return [];

    let afMonthKey = null;
    let cursor = previousMonthKey;
    let guard = 0;

    while (cursor && guard < 48) {
      const cursorKey = normalizeBillingMonthKeySafe_(cursor);
      if (!cursorKey) break;
      const flags = getBankWithdrawalStatusByPatient_(cursorKey, pid, cache);
      if (flags && flags.af) {
        afMonthKey = cursorKey;
        break;
      }
      cursor = resolvePreviousBillingMonthKey_(cursorKey);
      guard += 1;
    }

    if (afMonthKey) {
      if (!isNextMonth_(afMonthKey, monthKey)) return [];

      const aeMonths = [];
      cursor = resolvePreviousBillingMonthKey_(afMonthKey);
      guard = 0;

      while (cursor && guard < 48) {
        const cursorKey = normalizeBillingMonthKeySafe_(cursor);
        if (!cursorKey) break;
        const flags = getBankWithdrawalStatusByPatient_(cursorKey, pid, cache);
        if (flags && flags.ae) {
          aeMonths.unshift(cursorKey);
          cursor = resolvePreviousBillingMonthKey_(cursorKey);
          guard += 1;
          continue;
        }
        break;
      }

      return normalizePastBillingMonths_(aeMonths.concat(afMonthKey), monthKey);
    }

    return [previousMonthKey];
  })();

  RECEIPT_TARGET_MONTHS_BY_BANK_FLAGS_CACHE_[cacheKey] = result.slice();
  return result;
}

function isNextMonth_(fromMonthKey, toMonthKey) {
  const fromKey = normalizeBillingMonthKeySafe_(fromMonthKey);
  const toKey = normalizeBillingMonthKeySafe_(toMonthKey);
  if (!fromKey || !toKey) return false;
  const fromYear = Number(fromKey.slice(0, 4));
  const fromMonth = Number(fromKey.slice(4, 6));
  if (!fromYear || !fromMonth) return false;
  const nextMonth = fromMonth === 12 ? 1 : fromMonth + 1;
  const nextYear = fromMonth === 12 ? fromYear + 1 : fromYear;
  const nextKey = `${String(nextYear)}${String(nextMonth).padStart(2, '0')}`;
  return nextKey === toKey;
}

function resolveReceiptTargetMonths(patientId, billingMonth, cache) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  const pid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(patientId)
    : String(patientId || '').trim();
  if (!monthKey || !pid) return [];

  const previousMonthKey = resolvePreviousBillingMonthKey_(monthKey);
  if (!previousMonthKey) return [];

  const receiptMonths = resolveReceiptTargetMonthsByBankFlags_(patientId, monthKey, cache);
  if (!receiptMonths.length) return [];

  const previousPrepared = getPreparedBillingForMonthCached_(previousMonthKey, cache);
  const totals = previousPrepared && previousPrepared.totalsByPatient;
  if (!totals || !Object.prototype.hasOwnProperty.call(totals, pid)) return [];

  return receiptMonths;
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
    const aggregateTotals = resolveAggregateEntryTotalsForMonths_(targetMonths, pid, entry, monthCache);
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
      visitCount: aggregateTotals.visitCount,
      treatmentAmount: aggregateTotals.treatmentAmount,
      transportAmount: aggregateTotals.transportAmount,
      billingAmount: aggregateTotals.billingAmount,
      total: aggregateTotals.total,
      grandTotal: aggregateTotals.grandTotal,
      aggregateRemark,
      receiptMonths: targetMonths,
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
  const existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];
  const existingIndexes = [];
  existingRows.forEach((row, idx) => {
    if (String(row[0] || '').trim() === monthKey) {
      existingIndexes.push(idx);
    }
  });

  const normalizePid = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_
    : value => String(value || '').trim();

  const rows = billingJson.map(entry => {
    const rowPayload = entry || {};
    const pid = normalizePid(rowPayload.patientId);
    return [monthKey, pid || '', JSON.stringify(rowPayload)];
  });

  const updateCount = Math.min(existingIndexes.length, rows.length);
  for (let i = 0; i < updateCount; i++) {
    const rowIndex = existingIndexes[i] + 2;
    sheet.getRange(rowIndex, 1, 1, 3).setValues([rows[i]]);
  }

  const appendCount = rows.length - updateCount;
  if (appendCount > 0) {
    const appendStart = sheet.getLastRow() + 1;
    sheet.getRange(appendStart, 1, appendCount, 3).setValues(rows.slice(updateCount));
  }

  let cleared = 0;
  if (existingIndexes.length > rows.length) {
    for (let i = rows.length; i < existingIndexes.length; i++) {
      const rowIndex = existingIndexes[i] + 2;
      sheet.getRange(rowIndex, 1, 1, 3).setValues([['', '', '']]);
      cleared += 1;
    }
  }

  return { billingMonth: monthKey, inserted: rows.length, updated: updateCount, appended: appendCount, cleared };
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

  try {
    validatePreparedBillingPhase4A_(payload);
  } catch (err) {
    try {
      billingLogger_.log('[billing] Phase4-A validation failed to run: ' + err);
    } catch (logErr) {
      // ignore logging errors in non-GAS environments
    }
  }

  return { ok: true, billingMonth };
}

function logPreparedBillingPhase4AMismatch_(patientId, billingMonth, field, expected, actual) {
  try {
    billingLogger_.log('[billing] Phase4-A mismatch ' + JSON.stringify({
      patientId: patientId || '',
      month: billingMonth || '',
      field,
      expected,
      actual
    }));
  } catch (err) {
    try {
      console.warn('[billing] Phase4-A mismatch', { patientId, month: billingMonth, field, expected, actual });
    } catch (logErr) {
      // ignore logging errors in non-GAS environments
    }
  }
}

function validatePreparedBillingPhase4A_(payload) {
  if (!payload || typeof payload !== 'object') return;
  if (!Array.isArray(payload.billingJson) || payload.billingJson.length === 0) return;
  const billingMonth = payload.billingMonth || payload.month || '';
  const totalsByPatient = payload.totalsByPatient && typeof payload.totalsByPatient === 'object'
    ? payload.totalsByPatient
    : {};
  const normalizeAmount = typeof normalizeMoneyNumber_ === 'function'
    ? normalizeMoneyNumber_
    : value => Number(value) || 0;
  const isBlank = value => value === '' || value === null || value === undefined;
  const isSameAmount = (expected, actual) => {
    if (isBlank(expected) && isBlank(actual)) return true;
    const expectedNum = normalizeAmount(expected);
    const actualNum = normalizeAmount(actual);
    if (!Number.isFinite(expectedNum) || !Number.isFinite(actualNum)) {
      return String(expected || '') === String(actual || '');
    }
    return expectedNum === actualNum;
  };
  const compareAmount = (pid, field, expected, actual) => {
    if (isSameAmount(expected, actual)) return;
    logPreparedBillingPhase4AMismatch_(pid, billingMonth, field, expected, actual);
  };

  let bankAmounts = {};
  try {
    bankAmounts = collectBankWithdrawalAmountsByPatient_(billingMonth, payload) || {};
  } catch (err) {
    bankAmounts = {};
  }

  payload.billingJson.forEach(entry => {
    const pid = typeof billingNormalizePatientId_ === 'function'
      ? billingNormalizePatientId_(entry && entry.patientId)
      : String(entry && entry.patientId || '').trim();
    if (!pid) return;
    const insuranceEntry = resolveBillingEntryByType_(entry, 'insurance');
    const selfPayEntries = resolveBillingEntries_(entry).filter(item => (
      normalizeBillingEntryTypeValue_(item && (item.type || item.entryType)) === 'self_pay'
    ));
    const insuranceTotal = insuranceEntry ? resolveBillingEntryTotalAmount_(insuranceEntry) : 0;
    const selfPayTotal = selfPayEntries.reduce(
      (sum, item) => sum + resolveBillingEntryTotalAmount_(item),
      0
    );
    const expectedBillingAmount = insuranceEntry && Object.prototype.hasOwnProperty.call(insuranceEntry, 'billingAmount')
      ? normalizeAmount(insuranceEntry.billingAmount)
      : insuranceTotal;
    const expectedTotal = insuranceEntry ? insuranceTotal : 0;
    const expectedGrandTotal = insuranceTotal + selfPayTotal;

    compareAmount(pid, 'billingAmount', expectedBillingAmount, entry && entry.billingAmount);
    compareAmount(pid, 'total', expectedTotal, entry && entry.total);
    compareAmount(pid, 'grandTotal', expectedGrandTotal, entry && entry.grandTotal);

    const totalsEntry = totalsByPatient && totalsByPatient[pid];
    if (totalsEntry && typeof totalsEntry === 'object') {
      compareAmount(pid, 'totalsByPatient.billingAmount', expectedBillingAmount, totalsEntry.billingAmount);
      compareAmount(pid, 'totalsByPatient.total', expectedTotal, totalsEntry.total);
      compareAmount(pid, 'totalsByPatient.grandTotal', expectedGrandTotal, totalsEntry.grandTotal);
    }

    if (bankAmounts && Object.prototype.hasOwnProperty.call(bankAmounts, pid)) {
      compareAmount(pid, 'bankDebitAmount', insuranceTotal, bankAmounts[pid]);
    }
  });
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

function loadPreparedBillingForPdfGeneration_(billingMonthKey, options) {
  const opts = options || {};
  const withValidation = opts.withValidation === true;
  const allowInvalid = opts.allowInvalid === true;
  const wrapResult = (payload, validation) => withValidation ? { prepared: payload, validation: validation || null } : payload;

  const monthKey = normalizeBillingMonthKeySafe_(billingMonthKey);
  if (!monthKey) return wrapResult(null, { ok: false, reason: 'billingMonth missing' });

  const fromSheet = loadPreparedBillingFromSheet_(monthKey);
  const validation = validatePreparedBillingPayload_(fromSheet, monthKey);
  if (fromSheet && (validation.ok || allowInvalid)) {
    const preparedWithMonth = Object.assign({}, fromSheet, { billingMonth: validation.billingMonth || monthKey });
    return wrapResult(preparedWithMonth, validation);
  }

  return wrapResult(null, validation);
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

  const memoized = BILLING_CACHE_PAYLOAD_MEMO_[key];
  if (memoized && memoized.marker === cached && typeof memoized.payload === 'string') {
    return memoized.payload;
  }

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
  BILLING_CACHE_PAYLOAD_MEMO_[key] = { marker: cached, payload: merged };
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
  const patientMap = source.patients || source.patientMap || {};
  try {
    syncBankWithdrawalOnlineConsentFlags_(billingMonthKey || source.billingMonth || billingMonth, patientMap);
  } catch (err) {
    billingLogger_.log('[billing] Failed to sync online consent flags: ' + err);
  }
  const bankFlagsByPatient = getBankFlagsByPatient_(billingMonthKey || source.billingMonth || billingMonth, {
    patients: patientMap
  });
  source.bankFlagsByPatient = bankFlagsByPatient;
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
    const normalizedItem = normalizeBillingEntryFromEntries_(item);
    totalsByPatient[pid] = {
      visitCount: billingNormalizeVisitCount_(normalizedItem && normalizedItem.visitCount),
      billingAmount: normalizeNumber_(normalizedItem && normalizedItem.billingAmount),
      total: normalizeNumber_(normalizedItem && normalizedItem.total),
      grandTotal: normalizeNumber_(normalizedItem && normalizedItem.grandTotal),
      carryOverAmount: normalizeNumber_(normalizedItem && normalizedItem.carryOverAmount),
      carryOverFromHistory: normalizeNumber_(normalizedItem && normalizedItem.carryOverFromHistory)
    };
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
    billingOverridesSnapshot: source.billingOverridesSnapshot || '',
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
const BANK_WITHDRAWAL_SELF_PAY_COLUMN_LETTER = 'T';
const BANK_WITHDRAWAL_SELF_PAY_HEADER = '自費請求';
const BANK_WITHDRAWAL_DUPLICATE_WARNING_COLOR = '#fff2cc';

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
  ensureUnpaidCheckColumn_(copied, null, {
    billingJson: prepared && prepared.billingJson,
    bankRecords: prepared && prepared.bankRecords
  });
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
  const selfPayCol = resolveBillingColumn_(
    headers,
    [BANK_WITHDRAWAL_SELF_PAY_HEADER],
    BANK_WITHDRAWAL_SELF_PAY_HEADER,
    { fallbackLetter: BANK_WITHDRAWAL_SELF_PAY_COLUMN_LETTER }
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
  const selfPayAmountByPatientId = buildSelfPayAmountByPatientId_(prepared && prepared.billingJson);

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
  const selfPayValues = selfPayCol ? nameValues.map((row, idx) => {
    const rawPid = pidValues[idx] && pidValues[idx][0];
    const normalizedPid = typeof billingNormalizePatientId_ === 'function'
      ? billingNormalizePatientId_(rawPid)
      : (rawPid ? String(rawPid).trim() : '');
    const rawKana = (kanaValues[idx] && kanaValues[idx][0]) ? String(kanaValues[idx][0]).trim() : '';
    const fullNameKey = buildFullNameKey_(row && row[0], rawKana);
    const pid = normalizedPid || (fullNameKey ? nameToPatientId[fullNameKey] : '');
    const hasAmount = pid && Object.prototype.hasOwnProperty.call(selfPayAmountByPatientId, pid);
    const amount = hasAmount ? selfPayAmountByPatientId[pid] : null;
    if (!pid || amount === null || amount === undefined) {
      return [''];
    }
    return [amount];
  }) : [];

  copied.getRange(2, amountCol, rowCount, 1).setValues(amountValues);
  if (selfPayCol) {
    copied.getRange(2, selfPayCol, rowCount, 1).setValues(selfPayValues);
  }
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

  highlightDuplicateBankWithdrawalAccounts_(copied, headers, rowCount);

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

  const refreshHeaders_ = () => targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getDisplayValues()[0] || [];
  const insertColumnBefore_ = (index, label) => {
    targetSheet.insertColumnBefore(index);
    targetSheet.getRange(1, index).setValue(label);
  };

  let unpaidCol = resolveBillingColumn_(workingHeaders, [BANK_WITHDRAWAL_UNPAID_HEADER], BANK_WITHDRAWAL_UNPAID_HEADER, {});
  let aggregateCol = resolveBillingColumn_(workingHeaders, [BANK_WITHDRAWAL_AGGREGATE_HEADER], BANK_WITHDRAWAL_AGGREGATE_HEADER, {});
  let onlineCol = resolveBillingColumn_(workingHeaders, [BANK_WITHDRAWAL_ONLINE_HEADER], BANK_WITHDRAWAL_ONLINE_HEADER, {});
  if (!unpaidCol) {
    const insertBefore = [aggregateCol, onlineCol].filter(Boolean).sort((a, b) => a - b)[0];
    if (insertBefore) {
      insertColumnBefore_(insertBefore, BANK_WITHDRAWAL_UNPAID_HEADER);
      aggregateCol = aggregateCol ? aggregateCol + 1 : aggregateCol;
      onlineCol = onlineCol ? onlineCol + 1 : onlineCol;
      unpaidCol = insertBefore;
    } else {
      const lastCol = targetSheet.getLastColumn();
      targetSheet.insertColumnAfter(lastCol);
      unpaidCol = lastCol + 1;
      targetSheet.getRange(1, unpaidCol).setValue(BANK_WITHDRAWAL_UNPAID_HEADER);
    }
  }

  if (!aggregateCol) {
    const refreshedHeaders = refreshHeaders_();
    unpaidCol = resolveBillingColumn_(refreshedHeaders, [BANK_WITHDRAWAL_UNPAID_HEADER], BANK_WITHDRAWAL_UNPAID_HEADER, {});
    onlineCol = resolveBillingColumn_(refreshedHeaders, [BANK_WITHDRAWAL_ONLINE_HEADER], BANK_WITHDRAWAL_ONLINE_HEADER, {});
  }

  aggregateCol = aggregateCol || resolveBillingColumn_(refreshHeaders_(), [BANK_WITHDRAWAL_AGGREGATE_HEADER], BANK_WITHDRAWAL_AGGREGATE_HEADER, {});
  if (!aggregateCol) {
    try {
      targetSheet.insertColumnAfter(unpaidCol);
      aggregateCol = unpaidCol + 1;
      targetSheet.getRange(1, aggregateCol).setValue(BANK_WITHDRAWAL_AGGREGATE_HEADER);
    } catch (err) {
      console.warn('[billing] Failed to insert aggregate column on bank withdrawal sheet', err);
    }
  }

  onlineCol = resolveBillingColumn_(refreshHeaders_(), [BANK_WITHDRAWAL_ONLINE_HEADER], BANK_WITHDRAWAL_ONLINE_HEADER, {});
  if (!onlineCol) {
    try {
      const insertAfter = aggregateCol || unpaidCol || targetSheet.getLastColumn();
      targetSheet.insertColumnAfter(insertAfter);
      onlineCol = insertAfter + 1;
      targetSheet.getRange(1, onlineCol).setValue(BANK_WITHDRAWAL_ONLINE_HEADER);
    } catch (err) {
      console.warn('[billing] Failed to insert online column on bank withdrawal sheet', err);
    }
  }

  const alignBankWithdrawalFlagColumns_ = (unpaidIndex, aggregateIndex, onlineIndex) => {
    const indices = [unpaidIndex, aggregateIndex, onlineIndex].filter(Boolean);
    if (indices.length !== 3) return { unpaidCol: unpaidIndex, aggregateCol: aggregateIndex, onlineCol: onlineIndex };
    const targetPositions = indices.slice().sort((a, b) => a - b);
    const target = {
      unpaidCol: targetPositions[0],
      aggregateCol: targetPositions[1],
      onlineCol: targetPositions[2]
    };
    if (unpaidIndex === target.unpaidCol
      && aggregateIndex === target.aggregateCol
      && onlineIndex === target.onlineCol) {
      return { unpaidCol: unpaidIndex, aggregateCol: aggregateIndex, onlineCol: onlineIndex };
    }

    const lastRow = targetSheet.getLastRow();
    const rowCount = Math.max(lastRow, 1);
    const getColumnValues_ = columnIndex => targetSheet.getRange(1, columnIndex, rowCount, 1).getValues();
    const payload = {
      unpaid: getColumnValues_(unpaidIndex),
      aggregate: getColumnValues_(aggregateIndex),
      online: getColumnValues_(onlineIndex)
    };
    targetSheet.getRange(1, target.unpaidCol, rowCount, 1).setValues(payload.unpaid);
    targetSheet.getRange(1, target.aggregateCol, rowCount, 1).setValues(payload.aggregate);
    targetSheet.getRange(1, target.onlineCol, rowCount, 1).setValues(payload.online);

    targetSheet.getRange(1, target.unpaidCol).setValue(BANK_WITHDRAWAL_UNPAID_HEADER);
    targetSheet.getRange(1, target.aggregateCol).setValue(BANK_WITHDRAWAL_AGGREGATE_HEADER);
    targetSheet.getRange(1, target.onlineCol).setValue(BANK_WITHDRAWAL_ONLINE_HEADER);

    return target;
  };

  const aligned = alignBankWithdrawalFlagColumns_(unpaidCol, aggregateCol, onlineCol);
  unpaidCol = aligned.unpaidCol;
  aggregateCol = aligned.aggregateCol;
  onlineCol = aligned.onlineCol;

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
        if (onlineCol) {
          applyCheckboxes(onlineCol);
        }
      }
    }
  } catch (err) {
    console.warn('[billing] Failed to ensure checkbox columns for bank withdrawal flags', err);
  }

  enforceBankWithdrawalAggregateConstraint_(targetSheet, unpaidCol, aggregateCol);

  return { unpaidCol, aggregateCol, onlineCol };
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
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) {
    throw new Error('請求月が指定されていません');
  }
  const year = monthKey.slice(0, 4);
  const monthText = monthKey.slice(4, 6);
  return `${BANK_WITHDRAWAL_SHEET_PREFIX}${year}-${monthText}`;
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
  const aggregateCol = resolveBillingColumn_(headers, [BANK_WITHDRAWAL_AGGREGATE_HEADER], BANK_WITHDRAWAL_AGGREGATE_HEADER, {});
  const onlineCol = resolveBillingColumn_(headers, [BANK_WITHDRAWAL_ONLINE_HEADER], BANK_WITHDRAWAL_ONLINE_HEADER, {});
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

    const current = flagsByPatient[pid] || { ae: false, af: false, ag: false, online: false };
    const ae = unpaidCol ? normalizeBankFlagValue_(row[unpaidCol - 1]) : false;
    const af = aggregateCol ? normalizeBankFlagValue_(row[aggregateCol - 1]) : false;
    const ag = onlineCol ? normalizeBankFlagValue_(row[onlineCol - 1]) : false;
    flagsByPatient[pid] = {
      ae: current.ae || ae,
      af: current.af || af,
      ag: current.ag || ag,
      online: current.online || ag
    };
  });

  return flagsByPatient;
}

function syncBankWithdrawalOnlineConsentFlags_(billingMonth, prepared) {
  const month = normalizeBillingMonthInput(billingMonth);
  const patients = resolvePreparedPatients_(prepared);
  const patientIds = patients && typeof patients === 'object' ? Object.keys(patients) : [];
  if (!patientIds.length) return { billingMonth: month.key, updated: 0 };

  const sheet = ensureBankWithdrawalSheet_(month, { refreshFromTemplate: false, preserveExistingSheet: true });
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { billingMonth: month.key, updated: 0 };

  const headerCount = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, headerCount).getDisplayValues()[0];
  const flagCols = ensureBankWithdrawalFlagColumns_(sheet, headers, {});
  const onlineCol = flagCols && flagCols.onlineCol
    ? flagCols.onlineCol
    : resolveBillingColumn_(headers, [BANK_WITHDRAWAL_ONLINE_HEADER], BANK_WITHDRAWAL_ONLINE_HEADER, {});
  if (!onlineCol) return { billingMonth: month.key, updated: 0 };

  const pidCol = ensureBankWithdrawalPatientIdColumn_(sheet, headers);
  const refreshedHeaders = sheet.getRange(
    1,
    1,
    1,
    Math.max(sheet.getLastColumn(), headerCount, pidCol || 0, onlineCol)
  ).getDisplayValues()[0];
  const nameCol = resolveBillingColumn_(refreshedHeaders, BILLING_LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const kanaCol = resolveBillingColumn_(refreshedHeaders, BILLING_LABELS.furigana, 'フリガナ', {});
  const patientIdLabels = (typeof BILLING_LABELS !== 'undefined' && BILLING_LABELS && Array.isArray(BILLING_LABELS.recNo))
    ? BILLING_LABELS.recNo
    : [];
  const resolvedPidCol = pidCol || resolveBillingColumn_(refreshedHeaders, patientIdLabels.concat(['患者ID', '患者番号']), '患者ID', {});
  const rowCount = lastRow - 1;
  const nameValues = sheet.getRange(2, nameCol, rowCount, 1).getDisplayValues();
  const kanaValues = kanaCol ? sheet.getRange(2, kanaCol, rowCount, 1).getDisplayValues() : [];
  const pidValues = resolvedPidCol ? sheet.getRange(2, resolvedPidCol, rowCount, 1).getDisplayValues() : [];
  const onlineValues = sheet.getRange(2, onlineCol, rowCount, 1).getValues();
  const nameToPatientId = buildPatientNameToIdMap_(patients);
  const updatedValues = onlineValues.map(row => [row[0]]);

  let updated = 0;
  for (let idx = 0; idx < rowCount; idx++) {
    const rawPid = pidValues[idx] && pidValues[idx][0];
    const normalizedPid = typeof billingNormalizePatientId_ === 'function'
      ? billingNormalizePatientId_(rawPid)
      : (rawPid ? String(rawPid).trim() : '');
    const nameKey = buildFullNameKey_(
      nameValues[idx] && nameValues[idx][0],
      kanaValues[idx] && kanaValues[idx][0]
    );
    const pid = normalizedPid || (nameKey ? nameToPatientId[nameKey] : '');
    if (!pid || !patients || !patients[pid]) continue;
    const patient = patients[pid] || {};
    const consentFlag = typeof normalizeZeroOneFlag_ === 'function'
      ? normalizeZeroOneFlag_(patient.onlineConsent)
      : (patient.onlineConsent === true || patient.onlineConsent === 1 || patient.onlineConsent === '1' ? 1 : 0);
    const shouldBeOnline = !!consentFlag;
    const currentValue = updatedValues[idx][0];
    if (normalizeBankFlagValue_(currentValue) === shouldBeOnline) continue;
    updatedValues[idx][0] = shouldBeOnline;
    updated += 1;
  }

  if (updated) {
    sheet.getRange(2, onlineCol, rowCount, 1).setValues(updatedValues);
  }

  return { billingMonth: month.key, updated };
}

function resolveBillingEntryTotalAmount_(entry) {
  if (!entry) return 0;
  const manualOverride = entry.manualOverride && entry.manualOverride.amount;
  const normalizeAmount = typeof normalizeMoneyNumber_ === 'function'
    ? normalizeMoneyNumber_
    : value => Number(value) || 0;
  if (manualOverride !== '' && manualOverride !== null && manualOverride !== undefined) {
    return normalizeAmount(manualOverride);
  }
  return normalizeAmount(entry.total);
}

function normalizeBillingEntryTypeValue_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const normalized = raw.toLowerCase().replace(/\s+/g, '');
  if (normalized === 'insurance') return 'insurance';
  if (normalized === 'selfpay' || normalized === 'self_pay' || normalized === 'self-pay') return 'self_pay';
  return '';
}

function resolveBillingEntries_(entry) {
  if (!entry) return [];

  if (Array.isArray(entry.entries) && entry.entries.length) {
    return entry.entries
      .filter(
        item =>
          item &&
          typeof item === 'object' &&
          normalizeBillingEntryTypeValue_(item.type || item.entryType)
      )
      .map(item => {
        const normalizedType = normalizeBillingEntryTypeValue_(item.type || item.entryType);
        return Object.assign({}, item, {
          type: normalizedType,
          entryType: normalizedType === 'self_pay' ? 'selfPay' : normalizedType
        });
      });
  }

  return [];
}

function normalizeBillingEntryFromEntries_(entry) {
  if (!entry || typeof entry !== 'object') return entry;
  const normalizedEntry = Object.assign({}, entry);
  const normalizeAmount = typeof normalizeMoneyNumber_ === 'function'
    ? normalizeMoneyNumber_
    : value => Number(value) || 0;
  const normalizeVisits = typeof billingNormalizeVisitCount_ === 'function'
    ? billingNormalizeVisitCount_
    : value => Number(value) || 0;
  const buildFallbackEntries_ = source => {
    const fallbackEntries = [];
    const manualBillingInput = Object.prototype.hasOwnProperty.call(source, 'manualBillingAmount')
      ? source.manualBillingAmount
      : undefined;
    const hasManualBillingAmount =
      manualBillingInput !== '' && manualBillingInput !== null && manualBillingInput !== undefined;
    const insuranceBaseTotal = normalizeAmount(source.billingAmount)
      + normalizeAmount(source.transportAmount)
      + normalizeAmount(source.carryOverAmount)
      + normalizeAmount(source.carryOverFromHistory);
    const insuranceEntry = {
      type: 'insurance',
      entryType: 'insurance',
      unitPrice: source.unitPrice,
      visitCount: source.visitCount,
      treatmentAmount: source.treatmentAmount,
      transportAmount: source.transportAmount,
      billingAmount: source.billingAmount,
      total: hasManualBillingAmount ? normalizeAmount(manualBillingInput) : insuranceBaseTotal
    };
    if (hasManualBillingAmount) {
      insuranceEntry.manualOverride = { amount: normalizeAmount(manualBillingInput) };
    }
    fallbackEntries.push(insuranceEntry);

    const selfPayItems = Array.isArray(source.selfPayItems)
      ? source.selfPayItems.filter(item => item && typeof item === 'object')
      : [];
    const manualSelfPayInput = Object.prototype.hasOwnProperty.call(source, 'manualSelfPayAmount')
      ? source.manualSelfPayAmount
      : undefined;
    const hasManualSelfPayAmount =
      manualSelfPayInput !== '' && manualSelfPayInput !== null && manualSelfPayInput !== undefined;
    const selfPayItemsTotal = selfPayItems.reduce((sum, item) => sum + normalizeAmount(item.amount), 0);
    if (selfPayItems.length || hasManualSelfPayAmount) {
      const selfPayEntry = {
        type: 'self_pay',
        entryType: 'selfPay',
        items: selfPayItems,
        total: hasManualSelfPayAmount
          ? normalizeAmount(manualSelfPayInput)
          : selfPayItemsTotal
      };
      // Only include visitCount when visit-based self-pay is explicitly present.
      if (Object.prototype.hasOwnProperty.call(source, 'selfPayVisitCount')
        && source.selfPayVisitCount !== '' && source.selfPayVisitCount != null) {
        selfPayEntry.visitCount = normalizeVisits(source.selfPayVisitCount);
      }

      if (hasManualSelfPayAmount) {
        selfPayEntry.manualOverride = {
          amount: normalizeAmount(manualSelfPayInput)
        };
      }

      fallbackEntries.push(selfPayEntry);
    }

    return fallbackEntries;
  };
  let entries = resolveBillingEntries_(normalizedEntry);
  // Track fallback usage so we only backfill legacy row fields when we had to synthesize entries.
  const usedFallback = !entries.length;
  if (usedFallback) {
    entries = buildFallbackEntries_(normalizedEntry);
  }
  entries = entries.map(item => {
    if (!item || typeof item !== 'object') return item;
    const itemType = normalizeBillingEntryTypeValue_(item.type || item.entryType);
    const next = Object.assign({}, item);
    if (itemType === 'insurance') {
      if (usedFallback) {
        // Only backfill legacy row fields when entries were synthesized from legacy data.
        if (next.unitPrice == null && normalizedEntry.unitPrice != null) {
          next.unitPrice = normalizedEntry.unitPrice;
        }
        if (next.visitCount == null && normalizedEntry.visitCount != null) {
          next.visitCount = normalizeVisits(normalizedEntry.visitCount);
        }
        if (next.treatmentAmount == null && normalizedEntry.treatmentAmount != null) {
          next.treatmentAmount = normalizedEntry.treatmentAmount;
        }
        if (next.transportAmount == null && normalizedEntry.transportAmount != null) {
          next.transportAmount = normalizedEntry.transportAmount;
        }
        if (next.billingAmount == null && normalizedEntry.billingAmount != null) {
          next.billingAmount = normalizedEntry.billingAmount;
        }
        if (next.carryOverAmount == null && normalizedEntry.carryOverAmount != null) {
          next.carryOverAmount = normalizedEntry.carryOverAmount;
        }
        if (next.carryOverFromHistory == null && normalizedEntry.carryOverFromHistory != null) {
          next.carryOverFromHistory = normalizedEntry.carryOverFromHistory;
        }
        if (next.insuranceType == null && normalizedEntry.insuranceType != null) {
          next.insuranceType = normalizedEntry.insuranceType;
        }
        if (next.burdenRate == null && normalizedEntry.burdenRate != null) {
          next.burdenRate = normalizedEntry.burdenRate;
        }
        if (next.medicalAssistance == null && normalizedEntry.medicalAssistance != null) {
          next.medicalAssistance = normalizedEntry.medicalAssistance;
        }
        if (next.payerType == null && normalizedEntry.payerType != null) {
          next.payerType = normalizedEntry.payerType;
        }
        if (!Object.prototype.hasOwnProperty.call(next, 'manualUnitPrice')
          && Object.prototype.hasOwnProperty.call(normalizedEntry, 'manualUnitPrice')) {
          next.manualUnitPrice = normalizedEntry.manualUnitPrice;
        }
        if (!Object.prototype.hasOwnProperty.call(next, 'manualTransportAmount')
          && Object.prototype.hasOwnProperty.call(normalizedEntry, 'manualTransportAmount')) {
          next.manualTransportAmount = normalizedEntry.manualTransportAmount;
        }
        if (!Object.prototype.hasOwnProperty.call(next, 'adjustedVisitCount')
          && Object.prototype.hasOwnProperty.call(normalizedEntry, 'adjustedVisitCount')) {
          next.adjustedVisitCount = normalizedEntry.adjustedVisitCount;
        }
        if (!Object.prototype.hasOwnProperty.call(next, 'manualUnitPriceEntryType')
          && Object.prototype.hasOwnProperty.call(normalizedEntry, 'manualUnitPriceEntryType')) {
          next.manualUnitPriceEntryType = normalizedEntry.manualUnitPriceEntryType;
        }
        if (!Object.prototype.hasOwnProperty.call(next, 'manualBillingAmountEntryType')
          && Object.prototype.hasOwnProperty.call(normalizedEntry, 'manualBillingAmountEntryType')) {
          next.manualBillingAmountEntryType = normalizedEntry.manualBillingAmountEntryType;
        }
        if (!Object.prototype.hasOwnProperty.call(next, 'manualSelfPayAmountEntryType')
          && Object.prototype.hasOwnProperty.call(normalizedEntry, 'manualSelfPayAmountEntryType')) {
          next.manualSelfPayAmountEntryType = normalizedEntry.manualSelfPayAmountEntryType;
        }
      }
      return next;
    }
    if (itemType === 'self_pay') {
      if (usedFallback) {
        // Avoid injecting self-pay visit counts unless explicitly present in the entry.
        if (!Object.prototype.hasOwnProperty.call(next, 'manualSelfPayAmountEntryType')
          && Object.prototype.hasOwnProperty.call(normalizedEntry, 'manualSelfPayAmountEntryType')) {
          next.manualSelfPayAmountEntryType = normalizedEntry.manualSelfPayAmountEntryType;
        }
      }
      return next;
    }
    return next;
  });
  const resolveEntryByType_ = entryType => (
    entries.find(
      item =>
        item &&
        normalizeBillingEntryTypeValue_(item.type || item.entryType) ===
          normalizeBillingEntryTypeValue_(entryType)
    ) || null
  );
  const insuranceEntry = resolveEntryByType_('insurance');
  const selfPayEntries = entries.filter(
    item => item && normalizeBillingEntryTypeValue_(item.type || item.entryType) === 'self_pay'
  );
  const insuranceTotal = insuranceEntry ? resolveBillingEntryTotalAmount_(insuranceEntry) : 0;
  const selfPayTotal = selfPayEntries.reduce(
    (sum, item) => sum + resolveBillingEntryTotalAmount_(item),
    0
  );
  const expectedBillingAmount = insuranceEntry
    ? normalizeAmount(
      Object.prototype.hasOwnProperty.call(insuranceEntry, 'billingAmount')
        ? insuranceEntry.billingAmount
        : insuranceTotal
    )
    : 0;
  normalizedEntry.entries = entries;
  normalizedEntry.billingAmount = expectedBillingAmount;
  normalizedEntry.total = insuranceTotal;
  normalizedEntry.grandTotal = insuranceTotal + selfPayTotal;
  normalizedEntry.manualBillingAmount = insuranceEntry && insuranceEntry.manualOverride
    && Object.prototype.hasOwnProperty.call(insuranceEntry.manualOverride, 'amount')
    ? insuranceEntry.manualOverride.amount
    : '';
  const selfPayEntryWithManual = selfPayEntries.find(item => item && item.manualOverride
    && Object.prototype.hasOwnProperty.call(item.manualOverride, 'amount'));
  normalizedEntry.manualSelfPayAmount = selfPayEntryWithManual
    && Object.prototype.hasOwnProperty.call(selfPayEntryWithManual.manualOverride, 'amount')
    ? selfPayEntryWithManual.manualOverride.amount
    : '';
  return normalizedEntry;
}

function resolveBillingEntryByType_(entry, entryType) {
  const entries = resolveBillingEntries_(entry);
  const normalizedTarget = normalizeBillingEntryTypeValue_(entryType);

  return (
    entries.find(
      item =>
        item &&
        normalizeBillingEntryTypeValue_(item.type || item.entryType) ===
          normalizedTarget
    ) || null
  );
}

function buildBillingAmountByPatientId_(billingJson) {
  const amounts = {};
  (billingJson || []).forEach(entry => {
    const pid = billingNormalizePatientId_(entry && entry.patientId);
    if (!pid) return;
    const insuranceEntry = resolveBillingEntryByType_(entry, 'insurance');
    if (!insuranceEntry) return;
    amounts[pid] = resolveBillingEntryTotalAmount_(insuranceEntry);
  });
  return amounts;
}

function buildSelfPayAmountByPatientId_(billingJson) {
  const amounts = {};
  (billingJson || []).forEach(entry => {
    const pid = billingNormalizePatientId_(entry && entry.patientId);
    if (!pid) return;
    const selfPayEntries = resolveBillingEntries_(entry).filter(item => (
      normalizeBillingEntryTypeValue_(item && (item.type || item.entryType)) === 'self_pay'
    ));
    if (!selfPayEntries.length) return;
    amounts[pid] = selfPayEntries.reduce(
      (sum, item) => sum + resolveBillingEntryTotalAmount_(item),
      0
    );
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
  const selfPayCol = resolveBillingColumn_(
    headers,
    [BANK_WITHDRAWAL_SELF_PAY_HEADER],
    BANK_WITHDRAWAL_SELF_PAY_HEADER,
    { fallbackLetter: BANK_WITHDRAWAL_SELF_PAY_COLUMN_LETTER }
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
  const existingSelfPayValues = selfPayCol
    ? sheet.getRange(2, selfPayCol, rowCount, 1).getValues()
    : [];
  const patients = resolvePreparedPatients_(prepared);
  const nameToPatientId = buildPatientNameToIdMap_(patients);
  const amountByPatientId = buildBillingAmountByPatientId_(prepared && prepared.billingJson);
  const selfPayAmountByPatientId = buildSelfPayAmountByPatientId_(prepared && prepared.billingJson);
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
  const newSelfPayValues = selfPayCol ? nameValues.map((row, idx) => {
    const kanaRow = kanaValues[idx] || [];
    const rawPid = pidValues[idx] && pidValues[idx][0];
    const normalizedPid = typeof billingNormalizePatientId_ === 'function'
      ? billingNormalizePatientId_(rawPid)
      : (rawPid ? String(rawPid).trim() : '');
    const nameKey = buildFullNameKey_(row && row[0], kanaRow[0]);
    const pid = normalizedPid || (nameKey ? nameToPatientId[nameKey] : '');
    const resolvedAmount = pid && Object.prototype.hasOwnProperty.call(selfPayAmountByPatientId, pid)
      ? selfPayAmountByPatientId[pid]
      : null;
    const existingAmount = existingSelfPayValues[idx][0];
    const hasManualAmount = existingAmount !== '' && existingAmount !== null && existingAmount !== undefined;
    const nextAmount = hasManualAmount && resolvedAmount !== null && resolvedAmount !== undefined
      ? existingAmount
      : (resolvedAmount !== null && resolvedAmount !== undefined ? resolvedAmount : existingAmount);
    return [nextAmount];
  }) : [];

  const amountUpdates = [];
  newAmountValues.forEach((rowValue, idx) => {
    const existingValue = existingAmountValues[idx] ? existingAmountValues[idx][0] : '';
    const nextValue = rowValue[0];
    if (!isSameAmount_(existingValue, nextValue)) {
      amountUpdates.push({ row: idx + 2, value: nextValue });
    }
  });
  const selfPayUpdates = [];
  if (selfPayCol) {
    newSelfPayValues.forEach((rowValue, idx) => {
      const existingValue = existingSelfPayValues[idx] ? existingSelfPayValues[idx][0] : '';
      const nextValue = rowValue[0];
      if (!isSameAmount_(existingValue, nextValue)) {
        selfPayUpdates.push({ row: idx + 2, value: nextValue });
      }
    });
  }

  if (!amountUpdates.length && !selfPayUpdates.length) {
    highlightDuplicateBankWithdrawalAccounts_(sheet, headers, rowCount);
    return { billingMonth: month.key, updated: 0 };
  }

  if (amountUpdates.length) {
    let segmentStart = amountUpdates[0].row;
    let segmentValues = [[amountUpdates[0].value]];
    let previousRow = amountUpdates[0].row;
    for (let idx = 1; idx < amountUpdates.length; idx++) {
      const update = amountUpdates[idx];
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
  }

  if (selfPayCol && selfPayUpdates.length) {
    let segmentStart = selfPayUpdates[0].row;
    let segmentValues = [[selfPayUpdates[0].value]];
    let previousRow = selfPayUpdates[0].row;
    for (let idx = 1; idx < selfPayUpdates.length; idx++) {
      const update = selfPayUpdates[idx];
      if (update.row === previousRow + 1) {
        segmentValues.push([update.value]);
      } else {
        sheet.getRange(segmentStart, selfPayCol, segmentValues.length, 1).setValues(segmentValues);
        segmentStart = update.row;
        segmentValues = [[update.value]];
      }
      previousRow = update.row;
    }
    sheet.getRange(segmentStart, selfPayCol, segmentValues.length, 1).setValues(segmentValues);
  }

  highlightDuplicateBankWithdrawalAccounts_(sheet, headers, rowCount);

  return { billingMonth: month.key, updated: amountUpdates.length + selfPayUpdates.length };
}

function normalizeBankAccountKey_(value) {
  return String(value || '')
    .normalize('NFKC')
    .replace(/[\s\u3000・･．.ー－ｰ−‐-]+/g, '')
    .trim();
}

function highlightDuplicateBankWithdrawalAccounts_(sheet, headers, rowCount) {
  if (!sheet || !rowCount) return;
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return;

  const accountIdCol = resolveBillingColumn_(
    headers,
    ['口座ID', '口座Id', '口座ＩＤ', 'accountId', 'account_id'],
    '口座ID',
    {}
  );
  const accountNumberCol = resolveBillingColumn_(
    headers,
    ['口座番号', '口座No', '口座NO', 'accountNumber', '口座'],
    '口座番号',
    {}
  );

  if (!accountIdCol && !accountNumberCol) return;

  const accountIdValues = accountIdCol ? sheet.getRange(2, accountIdCol, rowCount, 1).getDisplayValues() : [];
  const accountNumberValues = accountNumberCol
    ? sheet.getRange(2, accountNumberCol, rowCount, 1).getDisplayValues()
    : [];

  const rowsByKey = {};
  for (let idx = 0; idx < rowCount; idx++) {
    const rawAccountId = accountIdValues[idx] ? accountIdValues[idx][0] : '';
    const rawAccountNumber = accountNumberValues[idx] ? accountNumberValues[idx][0] : '';
    const normalizedAccountId = normalizeBankAccountKey_(rawAccountId);
    const normalizedAccountNumber = normalizeBankAccountKey_(rawAccountNumber);
    const key = normalizedAccountId || normalizedAccountNumber;
    if (!key) continue;
    if (!rowsByKey[key]) rowsByKey[key] = [];
    rowsByKey[key].push(idx);
  }

  const duplicateRows = new Set();
  Object.keys(rowsByKey).forEach(key => {
    const indices = rowsByKey[key];
    if (indices && indices.length > 1) {
      indices.forEach(idx => duplicateRows.add(idx));
    }
  });

  const range = sheet.getRange(2, 1, rowCount, lastCol);
  const backgrounds = range.getBackgrounds();
  let needsUpdate = false;

  for (let rowIdx = 0; rowIdx < rowCount; rowIdx++) {
    if (duplicateRows.has(rowIdx)) {
      for (let colIdx = 0; colIdx < lastCol; colIdx++) {
        if (backgrounds[rowIdx][colIdx] !== BANK_WITHDRAWAL_DUPLICATE_WARNING_COLOR) {
          backgrounds[rowIdx][colIdx] = BANK_WITHDRAWAL_DUPLICATE_WARNING_COLOR;
          needsUpdate = true;
        }
      }
    } else {
      const isAllWarning = backgrounds[rowIdx].every(
        color => color === BANK_WITHDRAWAL_DUPLICATE_WARNING_COLOR
      );
      if (isAllWarning) {
        for (let colIdx = 0; colIdx < lastCol; colIdx++) {
          backgrounds[rowIdx][colIdx] = '';
        }
        needsUpdate = true;
      }
    }
  }

  if (needsUpdate) {
    range.setBackgrounds(backgrounds);
  }
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
      const sanitized = sanitizeAggregateFieldsForBankFlags_(normalizedEntry, normalizedEntry.bankFlags);
      return normalizeBillingEntryFromEntries_(sanitized);
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

function resetPreparedBillingAndPrepare(billingMonth, options) {
  const normalizedMonth = normalizeBillingMonthInput ? normalizeBillingMonthInput(billingMonth) : null;
  const monthKey = normalizedMonth && normalizedMonth.key
    ? normalizedMonth.key
    : normalizeBillingMonthKeySafe_(billingMonth);
  if (!monthKey) {
    throw new Error('請求月をYYYY-MM形式で指定してください。');
  }

  try {
    if (typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function') {
      billingLogger_.log('[billing] resetPreparedBillingAndPrepare deleting prepared billing for ' + monthKey);
    }
  } catch (err) {
    // ignore logging issues in non-GAS environments
  }

  deletePreparedBillingDataForMonth_(monthKey);
  const nextOptions = Object.assign({}, options || {}, { forceReaggregate: true });
  return prepareBillingData(monthKey, nextOptions);
}

function normalizeOnlineConsentFlag_(value) {
  if (typeof normalizeZeroOneFlag_ === 'function') {
    return normalizeZeroOneFlag_(value);
  }
  return value === true || value === 1 || value === '1' ? 1 : 0;
}

function hasOnlineConsentMismatch_(currentPatients, existingPatients) {
  const currentKeys = Object.keys(currentPatients || {});
  for (let i = 0; i < currentKeys.length; i++) {
    const pid = currentKeys[i];
    const current = currentPatients && currentPatients[pid];
    const existing = existingPatients && existingPatients[pid];
    const currentFlag = normalizeOnlineConsentFlag_(current && current.onlineConsent);
    const existingFlag = normalizeOnlineConsentFlag_(existing && existing.onlineConsent);
    if (currentFlag !== existingFlag) return true;
  }
  return false;
}

function hasOnlineBankFlagMismatch_(currentFlags, existingFlags) {
  const allKeys = Object.keys(currentFlags || {}).concat(Object.keys(existingFlags || {}));
  const unique = Array.from(new Set(allKeys));
  for (let i = 0; i < unique.length; i++) {
    const pid = unique[i];
    const currentFlag = normalizeBankFlagValue_(currentFlags && currentFlags[pid] && currentFlags[pid].ag);
    const existingFlag = normalizeBankFlagValue_(existingFlags && existingFlags[pid] && existingFlags[pid].ag);
    if (currentFlag !== existingFlag) return true;
  }
  return false;
}

function shouldInvalidatePreparedBilling_(existingPrepared, source, monthKey) {
  if (!existingPrepared || !source) return false;
  const patients = source.patients || source.patientMap || {};
  const existingPatients = existingPrepared.patients || {};
  if (hasOnlineConsentMismatch_(patients, existingPatients)) {
    billingLogger_.log('[billing] prepareBillingData invalidate: onlineConsent mismatch for ' + monthKey);
    return true;
  }
  const currentFlags = source.bankFlagsByPatient || {};
  const existingFlags = existingPrepared.bankFlagsByPatient || {};
  if (hasOnlineBankFlagMismatch_(currentFlags, existingFlags)) {
    billingLogger_.log('[billing] prepareBillingData invalidate: bankFlags mismatch for ' + monthKey);
    return true;
  }
  const currentSnapshot = source.billingOverridesSnapshot || '';
  const existingSnapshot = existingPrepared.billingOverridesSnapshot || '';
  if (currentSnapshot !== existingSnapshot) {
    billingLogger_.log('[billing] prepareBillingData invalidate: billing overrides mismatch for ' + monthKey);
    return true;
  }
  return false;
}

function prepareBillingData(billingMonth, options) {
  const normalizedMonth = normalizeBillingMonthInput(billingMonth);
  let shouldForceReaggregate = options && options.forceReaggregate === true;
  if (!shouldForceReaggregate) {
    const existingResult = loadPreparedBillingWithSheetFallback_(normalizedMonth.key, { withValidation: true });
    const existingPrepared = existingResult && existingResult.prepared !== undefined ? existingResult.prepared : existingResult;
    const existingValidation = existingResult && existingResult.validation
      ? existingResult.validation
      : (existingResult && existingResult.ok !== undefined ? existingResult : null);
    if (existingPrepared && (!existingValidation || existingValidation.ok)) {
      const currentSource = getBillingSourceData(normalizedMonth.key);
      const currentPatients = currentSource.patients || currentSource.patientMap || {};
      try {
        syncBankWithdrawalOnlineConsentFlags_(normalizedMonth.key, currentPatients);
      } catch (err) {
        billingLogger_.log('[billing] prepareBillingData failed to sync online consent flags: ' + err);
      }
      const currentBankFlags = getBankFlagsByPatient_(normalizedMonth.key, { patients: currentPatients });
      currentSource.bankFlagsByPatient = currentBankFlags;
      if (shouldInvalidatePreparedBilling_(existingPrepared, currentSource, normalizedMonth.key)) {
        shouldForceReaggregate = true;
      }
    }
    if (shouldForceReaggregate) {
      billingLogger_.log('[billing] prepareBillingData invalidated prepared billing for ' + normalizedMonth.key);
    } else if (existingPrepared && (!existingValidation || existingValidation.ok)) {
      const manualSync = syncManualBillingOverridesIntoPrepared_(existingPrepared, normalizedMonth.key);
      if (manualSync && manualSync.updated) {
        savePreparedBillingToSheet_(normalizedMonth.key, manualSync.prepared);
        billingLogger_.log('[billing] prepareBillingData refreshed manual billing overrides for ' + normalizedMonth.key);
      }
      if (manualSync && manualSync.prepared) {
        savePreparedBilling_(manualSync.prepared);
        if (!manualSync.updated) {
          billingLogger_.log('[billing] prepareBillingData using existing prepared billing for ' + normalizedMonth.key);
        }
        return manualSync.prepared;
      }
      billingLogger_.log('[billing] prepareBillingData using existing prepared billing for ' + normalizedMonth.key);
      return existingPrepared;
    }
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
  const shouldSyncBank = !(options && options.syncBankWithdrawal === false);
  let bankSheetResult = null;
  if (shouldSyncBank) {
    bankSheetResult = syncBankWithdrawalSheetForMonth_(normalizedMonth, normalizedPrepared);
    billingLogger_.log('[billing] Bank withdrawal sheet synced: ' + JSON.stringify(bankSheetResult));
  } else {
    billingLogger_.log('[billing] Bank withdrawal sheet sync skipped for ' + normalizedMonth.key);
  }
  return cachePayload;
}

function syncManualBillingOverridesIntoPrepared_(prepared, billingMonth) {
  const monthKey = normalizeBillingMonthKeySafe_(billingMonth || (prepared && prepared.billingMonth));
  if (!monthKey || !prepared || !Array.isArray(prepared.billingJson)) {
    return { prepared, updated: false };
  }
  if (typeof loadBillingOverridesMap_ !== 'function') {
    return { prepared, updated: false };
  }
  const overrides = loadBillingOverridesMap_(monthKey);
  if (!overrides || !Object.keys(overrides).length) {
    return { prepared, updated: false };
  }

  let updated = false;
  const updatedBillingJson = prepared.billingJson.map(entry => {
    if (!entry) return entry;
    const pid = billingNormalizePatientId_(entry.patientId);
    if (!pid) return entry;
    const override = overrides[pid];
    if (!override || (override.manualBillingAmount === undefined && override.manualSelfPayAmount === undefined)) {
      return normalizeBillingEntryFromEntries_(entry);
    }
    const current = normalizeBillingEntryFromEntries_(entry);
    const entries = resolveBillingEntries_(current);
    const insuranceOverride = override.manualBillingAmount;
    const selfPayOverride = override.manualSelfPayAmount;
    const insuranceOverrideEntryType = normalizeBillingEntryType_(override.manualBillingAmountEntryType || override.entryType);
    const selfPayOverrideEntryType = normalizeBillingEntryType_(override.manualSelfPayAmountEntryType || override.entryType);
    const updatedEntries = entries.map(item => {
      const itemType = normalizeBillingEntryType_(item && (item.type || item.entryType));
      if (itemType === 'insurance'
        && insuranceOverride !== undefined
        && (!insuranceOverrideEntryType || insuranceOverrideEntryType === itemType)) {
        const next = Object.assign({}, item);
        if (insuranceOverride === '' || insuranceOverride === null) {
          if (next.manualOverride && Object.prototype.hasOwnProperty.call(next.manualOverride, 'amount')) {
            const cleaned = Object.assign({}, next.manualOverride);
            delete cleaned.amount;
            if (!Object.keys(cleaned).length) {
              delete next.manualOverride;
            } else {
              next.manualOverride = cleaned;
            }
          }
        } else {
          next.manualOverride = Object.assign({}, next.manualOverride || {}, {
            amount: insuranceOverride
          });
        }
        return next;
      }
      if (itemType === 'self_pay'
        && selfPayOverride !== undefined
        && (!selfPayOverrideEntryType || selfPayOverrideEntryType === itemType)) {
        const next = Object.assign({}, item);
        if (selfPayOverride === '' || selfPayOverride === null) {
          if (next.manualOverride && Object.prototype.hasOwnProperty.call(next.manualOverride, 'amount')) {
            const cleaned = Object.assign({}, next.manualOverride);
            delete cleaned.amount;
            if (!Object.keys(cleaned).length) {
              delete next.manualOverride;
            } else {
              next.manualOverride = cleaned;
            }
          }
        } else {
          next.manualOverride = Object.assign({}, next.manualOverride || {}, {
            amount: selfPayOverride
          });
        }
        return next;
      }
      return item;
    });
    const normalized = normalizeBillingEntryFromEntries_(Object.assign({}, current, { entries: updatedEntries }));
    const billingChanged = override.manualBillingAmount !== undefined
      && normalized.manualBillingAmount !== current.manualBillingAmount;
    const selfPayChanged = override.manualSelfPayAmount !== undefined
      && normalized.manualSelfPayAmount !== current.manualSelfPayAmount;
    if (billingChanged || selfPayChanged) {
      updated = true;
    }
    return normalized;
  });

  if (!updated) {
    return { prepared, updated: false };
  }

  const updatedPrepared = Object.assign({}, prepared, { billingJson: updatedBillingJson });
  if (updatedPrepared.totalsByPatient && typeof updatedPrepared.totalsByPatient === 'object') {
    const totalsByPatient = Object.assign({}, updatedPrepared.totalsByPatient);
    updatedBillingJson.forEach(entry => {
      if (!entry) return;
      const pid = billingNormalizePatientId_(entry.patientId);
      if (!pid || !totalsByPatient[pid]) return;
      totalsByPatient[pid] = Object.assign({}, totalsByPatient[pid], {
        grandTotal: normalizeMoneyNumber_(entry.grandTotal),
        total: normalizeMoneyNumber_(entry.total),
        billingAmount: normalizeMoneyNumber_(entry.billingAmount)
      });
    });
    updatedPrepared.totalsByPatient = totalsByPatient;
  }
  if (updatedPrepared.billingOverrideFlags && typeof updatedPrepared.billingOverrideFlags === 'object') {
    const overrideFlags = Object.assign({}, updatedPrepared.billingOverrideFlags);
    Object.keys(overrides).forEach(pid => {
      const override = overrides[pid];
      if (!override) return;
      const flags = Object.assign({}, overrideFlags[pid] || {});
      if (override.manualBillingAmount !== undefined
        && override.manualBillingAmount !== '' && override.manualBillingAmount != null) {
        flags.grandTotal = true;
      }
      if (override.manualSelfPayAmount !== undefined
        && override.manualSelfPayAmount !== '' && override.manualSelfPayAmount != null) {
        flags.selfPayAmount = true;
      }
      overrideFlags[pid] = flags;
    });
    updatedPrepared.billingOverrideFlags = overrideFlags;
  }
  return { prepared: updatedPrepared, updated: true };
}

/**
 * Generate billing JSON without creating files (preview use case).
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @return {Object} billingMonth key and generated JSON array.
 */
function generateBillingJsonPreview(billingMonth) {
  const prepared = buildPreparedBillingPayload_(billingMonth);
  const normalized = normalizePreparedBilling_(prepared) || prepared;
  savePreparedBilling_(normalized);
  savePreparedBillingToSheet_(billingMonth, normalized);
  return normalized;
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
    const normalized = normalizePreparedBilling_(prepared) || prepared;
    savePreparedBilling_(normalized);
    savePreparedBillingToSheet_(billingMonth, normalized);
    const monthCache = createBillingMonthCache_();
    loadPreparedBillingSummariesIntoCache_(monthCache);
    loadBankWithdrawalAmountsIntoCache_(monthCache, normalized);
    return generatePreparedInvoices_(normalized, Object.assign({}, options || {}, { monthCache }));
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

function shouldGenerateInvoicePdfForEntry_(entry, prepared) {
  if (!entry) return false;
  const patientId = billingNormalizePatientId_(entry && entry.patientId);
  if (!patientId) return false;

  const patientMap = prepared && prepared.patients ? prepared.patients : null;
  const patient = patientMap && Object.prototype.hasOwnProperty.call(patientMap, patientId)
    ? patientMap[patientId]
    : null;
  if (patient && typeof normalizeMedicalSubsidyFlag_ === 'function') {
    const hasMedicalSubsidy = normalizeMedicalSubsidyFlag_(patient.medicalSubsidy);
    if (hasMedicalSubsidy) return false;
  }

  const insuranceEntry = resolveBillingEntryByType_(entry, 'insurance');
  if (!insuranceEntry) return false;
  const normalizeVisits = typeof billingNormalizeVisitCount_ === 'function'
    ? billingNormalizeVisitCount_
    : value => Number(value) || 0;
  const visitCount = normalizeVisits(insuranceEntry && insuranceEntry.visitCount);
  if (!visitCount) return false;

  const billingAmount = normalizeMoneyNumber_(insuranceEntry && insuranceEntry.billingAmount != null
    ? insuranceEntry.billingAmount
    : insuranceEntry.total);
  const grandTotal = normalizeMoneyNumber_(resolveBillingEntryTotalAmount_(insuranceEntry));
  if (!billingAmount || !grandTotal) return false;

  return true;
}

function shouldGenerateSelfPayInvoicePdfForEntry_(entry) {
  if (!entry) return false;
  const selfPayEntries = resolveBillingEntries_(entry).filter(item => (
    normalizeBillingEntryTypeValue_(item && (item.type || item.entryType)) === 'self_pay'
  ));
  if (!selfPayEntries.length) return false;
  const selfPayItems = buildSelfPayItemsForInvoice_(entry);
  if (selfPayItems.length) return true;
  return selfPayEntries.some(item => resolveBillingEntryTotalAmount_(item) > 0);
}

function buildSelfPayItemsForInvoice_(entry) {
  const selfPayEntries = resolveBillingEntries_(entry).filter(item => (
    normalizeBillingEntryTypeValue_(item && (item.type || item.entryType)) === 'self_pay'
  ));
  if (!selfPayEntries.length) return [];
  const rawItems = selfPayEntries.reduce((list, entryItem) => {
    const items = Array.isArray(entryItem.items)
      ? entryItem.items
      : (Array.isArray(entryItem.selfPayItems) ? entryItem.selfPayItems : []);
    return list.concat(items);
  }, []);
  const items = rawItems.filter(item => item && normalizeMoneyNumber_(item.amount) !== 0);
  const manualOverrideEntry = selfPayEntries.find(entryItem => entryItem && entryItem.manualOverride
    && entryItem.manualOverride.amount !== '' && entryItem.manualOverride.amount !== null
    && entryItem.manualOverride.amount !== undefined);
  const manualOverride = manualOverrideEntry ? manualOverrideEntry.manualOverride.amount : undefined;
  if (!items.length && manualOverride != null && manualOverride !== '') {
    const manualAmount = normalizeMoneyNumber_(manualOverride);
    if (manualAmount !== 0) {
      items.push({ type: '自費', amount: manualAmount });
    }
  }
  return items;
}

function buildSelfPayInvoiceEntryForPdf_(entry) {
  if (!entry) return null;
  const selfPayEntries = resolveBillingEntries_(entry).filter(item => (
    normalizeBillingEntryTypeValue_(item && (item.type || item.entryType)) === 'self_pay'
  ));
  if (!selfPayEntries.length) return null;
  const selfPayCount = Number(entry.selfPayCount) || 0;
  const selfPayItems = buildSelfPayItemsForInvoice_(entry);
  const manualOverrideEntry = selfPayEntries.find(entryItem => entryItem && entryItem.manualOverride
    && entryItem.manualOverride.amount !== '' && entryItem.manualOverride.amount !== null
    && entryItem.manualOverride.amount !== undefined);
  const manualOverride = manualOverrideEntry ? manualOverrideEntry.manualOverride.amount : undefined;
  return Object.assign({}, entry, {
    unitPrice: 0,
    visitCount: selfPayCount,
    insuranceType: '自費',
    burdenRate: '自費',
    selfPayItems,
    manualSelfPayAmount: manualOverride !== undefined ? manualOverride : '',
    carryOverAmount: 0,
    manualTransportAmount: '',
    transportAmount: 0,
    treatmentAmount: 0,
    billingAmount: null,
    total: null,
    grandTotal: null
  });
}

function buildInsuranceInvoiceEntryForPdf_(entry) {
  if (!entry) return null;
  const insuranceEntry = resolveBillingEntryByType_(entry, 'insurance');
  if (!insuranceEntry) return null;
  const manualOverride = insuranceEntry && insuranceEntry.manualOverride
    ? insuranceEntry.manualOverride.amount
    : undefined;
  const insuranceTotal = resolveBillingEntryTotalAmount_(insuranceEntry);
  const selfPayItems = buildSelfPayItemsForInvoice_(entry);
  return Object.assign({}, entry, {
    visitCount: insuranceEntry.visitCount,
    unitPrice: insuranceEntry.unitPrice,
    treatmentAmount: insuranceEntry.treatmentAmount,
    transportAmount: insuranceEntry.transportAmount,
    billingAmount: insuranceEntry.billingAmount,
    total: insuranceTotal,
    grandTotal: insuranceTotal,
    manualBillingAmount: manualOverride !== undefined ? manualOverride : '',
    manualSelfPayAmount: '',
    selfPayItems,
    selfPayCount: Number(entry && entry.selfPayCount) || 0
  });
}

function buildSelfPayInvoicePdfContextForEntry_(entry, prepared, cache) {
  if (!entry) return null;
  const billingMonth = normalizeBillingMonthKeySafe_(prepared && prepared.billingMonth ? prepared.billingMonth : entry.billingMonth);
  const patientId = billingNormalizePatientId_(entry && entry.patientId);
  if (!billingMonth || !patientId) return null;
  const amount = finalizeInvoiceAmountDataForPdf_(entry, billingMonth, [], false, cache, prepared);
  return {
    patientId,
    billingMonth,
    months: [],
    amount,
    name: entry && entry.nameKanji ? String(entry.nameKanji) : '',
    isAggregateInvoice: false,
    responsibleName: entry && entry.responsibleName ? String(entry.responsibleName) : ''
  };
}

function appendSelfPaySuffixToFileName_(fileName) {
  const name = String(fileName || '').trim();
  if (!name) return '';
  if (name.indexOf('_自費') >= 0) return name;
  const lowerName = name.toLowerCase();
  if (lowerName.endsWith('.pdf')) {
    return name.slice(0, -4) + '_自費' + name.slice(-4);
  }
  return name + '_自費';
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
  const selfPayItems = Array.isArray(entry && entry.selfPayItems)
    ? entry.selfPayItems
    : [];
  selfPayItems.forEach(item => {
    if (!item) return;
    rows.push({ label: resolveInvoiceItemLabel_(item), detail: '', amount: item.amount });
  });

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

function finalizeInvoiceAmountDataForPdf_(entry, billingMonth, aggregateMonths, isAggregateInvoice, cache, prepared) {
  const normalizedAggregateMonths = normalizePastBillingMonths_(aggregateMonths, billingMonth);
  const baseAmount = isAggregateInvoice
    ? buildAggregateInvoiceAmountDataForPdf_(normalizedAggregateMonths, billingMonth, entry && entry.patientId, cache)
    : buildStandardInvoiceAmountDataForPdf_(entry, billingMonth);
  const displayAmount = buildStandardInvoiceAmountDataForPdf_(entry, billingMonth);
  const aggregateMonthTotals = Array.isArray(baseAmount.aggregateMonthTotals)
    ? baseAmount.aggregateMonthTotals
    : [];
  const receiptMonthBreakdown = Array.isArray(entry && entry.receiptMonthBreakdown) ? entry.receiptMonthBreakdown : [];
  const breakdownTotal = receiptMonthBreakdown.length
    ? receiptMonthBreakdown.reduce((sum, item) => sum + (normalizeMoneyNumber_(item && item.amount) || 0), 0)
    : null;
  const hasBreakdownAmount = receiptMonthBreakdown.some(item => (normalizeMoneyNumber_(item && item.amount) || 0) > 0);
  const normalizedPreviousReceiptAmount = entry && entry.previousReceiptAmount != null
    ? normalizeMoneyNumber_(entry.previousReceiptAmount)
    : undefined;
  const resolvedPreviousReceiptAmount = hasBreakdownAmount
    ? breakdownTotal
    : (normalizedPreviousReceiptAmount != null ? normalizedPreviousReceiptAmount : undefined);
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
  if (entry && entry.grandTotal != null && entry.grandTotal !== '') {
    amount.grandTotal = normalizeMoneyNumber_(entry.grandTotal);
  }
  if (isAggregateInvoice && aggregateMonthTotals.length) {
    amount.grandTotal = aggregateMonthTotals.reduce(
      (sum, row) => sum + (normalizeMoneyNumber_(row && row.total) || 0),
      0
    );
  }

  const receiptDisplay = resolveInvoiceReceiptDisplay_(entry, { aggregateMonths: normalizedAggregateMonths });
  const basePreviousReceipt = buildInvoicePreviousReceipt_(entry, receiptDisplay, normalizedAggregateMonths);
  const previousReceipt = entry && entry.previousReceipt
    ? Object.assign({}, basePreviousReceipt, entry.previousReceipt)
    : basePreviousReceipt;
  if (previousReceipt) {
    previousReceipt.settled = isPreviousReceiptSettled_(entry);
    if (resolvedPreviousReceiptAmount != null) {
      previousReceipt.amount = resolvedPreviousReceiptAmount;
    }
  }

  const aggregateStatus = receiptDisplay ? receiptDisplay.aggregateStatus : normalizeAggregateStatus_(entry && entry.aggregateStatus);
  const aggregateConfirmed = !!(receiptDisplay && receiptDisplay.aggregateConfirmed);
  const watermark = buildInvoiceWatermark_(entry);
  const receiptMonths = resolveReceiptTargetMonths(entry && entry.patientId, billingMonth, cache);
  const currentFlags = getBankWithdrawalStatusByPatient_(billingMonth, entry && entry.patientId, cache);
  const isAggregateMonth = !!(currentFlags && currentFlags.af);
  let shouldShowReceipt = !!(receiptDisplay && receiptDisplay.visible);
  const carryOverAmount = normalizeBillingCarryOver_(entry);
  const displayFlags = resolveInvoiceDisplayMode_(entry, billingMonth, {
    showReceipt: shouldShowReceipt,
    previousReceipt,
    previousReceiptAmount: resolvedPreviousReceiptAmount,
    carryOverAmount
  });
  if (displayFlags.displayMode === 'aggregate' && isAggregateMonth) {
    shouldShowReceipt = false;
    amount.forceHideReceipt = true;
  }
  if (previousReceipt) {
    previousReceipt.visible = shouldShowReceipt;
  }
  const entrySelfPayItems = Array.isArray(entry && entry.selfPayItems)
    ? entry.selfPayItems
    : [];
  const previousReceiptMonth = Array.isArray(receiptDisplay && receiptDisplay.receiptMonths) && receiptDisplay.receiptMonths.length
    ? receiptDisplay.receiptMonths[receiptDisplay.receiptMonths.length - 1]
    : '';
  logReceiptDebug_(entry && entry.patientId, {
    step: 'finalizeInvoiceAmountDataForPdf_',
    billingMonth,
    patientId: entry && entry.patientId,
    receiptTargetMonths: normalizedAggregateMonths,
    receiptMonths
  });
  logReceiptDebug_(entry && entry.patientId, {
    step: 'finalizeInvoiceAmountDataForPdf_ pre-return',
    billingMonth,
    patientId: entry && entry.patientId,
    isAggregateInvoice,
    carryOverAmount,
    previousReceipt,
    previousReceiptAmount: entry && entry.previousReceiptAmount != null ? entry.previousReceiptAmount : undefined,
    receiptMonths,
    receiptDisplayVisible: !!(receiptDisplay && receiptDisplay.visible),
    aggregateStatus,
    aggregateConfirmed
  });

  return Object.assign({}, amount, {
    aggregateStatus,
    aggregateConfirmed,
    receiptMonths,
    receiptRemark: receiptDisplay && receiptDisplay.receiptRemark ? receiptDisplay.receiptRemark : '',
    receiptMonthBreakdown: receiptMonthBreakdown.length ? receiptMonthBreakdown : undefined,
    previousReceiptAmount: resolvedPreviousReceiptAmount,
    previousReceiptMonth,
    selfPayItems: displayFlags.displayMode === 'standard' ? entrySelfPayItems : [],
    showReceipt: shouldShowReceipt,
    previousReceipt,
    displayMode: displayFlags.displayMode,
    showOnlineConsentNote: displayFlags.showOnlineConsentNote,
    showPreviousReceipt: displayFlags.showPreviousReceipt,
    displayRows: Array.isArray(displayAmount && displayAmount.rows) ? displayAmount.rows : [],
    carryOverAmount,
    watermark
  });
}

function buildInvoicePdfContextForEntry_(entry, prepared, cache, receiptSummaryMap) {
  if (!entry) return null;
  const billingMonth = normalizeBillingMonthKeySafe_(prepared && prepared.billingMonth ? prepared.billingMonth : entry.billingMonth);
  const patientId = billingNormalizePatientId_(entry && entry.patientId);
  if (!billingMonth || !patientId) return null;

  const receiptEntry = entry;
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
  const receiptSummary = resolveReceiptSummaryForPatient_(patientId, billingMonth, cache, receiptSummaryMap);
  const aggregateMonths = receiptSummary.aggregateMonths || [];
  const isAggregateInvoice = !!receiptSummary.isAggregateInvoice;
  const amount = finalizeInvoiceAmountDataForPdf_(receiptEntry, billingMonth, aggregateMonths, isAggregateInvoice, cache, prepared);

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
  const amount = finalizeInvoiceAmountDataForPdf_(entry, billingMonth, normalizedMonths, true, cache, prepared);

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
  const opts = options || {};
  const includeInsurancePdf = opts.includeInsurancePdf !== false;
  const includeSelfPayPdf = opts.includeSelfPayPdf === true;
  const monthCache = opts.monthCache || createBillingMonthCache_();
  if (normalized && normalized.billingMonth && monthCache.preparedByMonth) {
    monthCache.preparedByMonth[normalized.billingMonth] = monthCache.preparedByMonth[normalized.billingMonth]
      || reducePreparedBillingSummary_(normalized);
  }
  preloadPreparedPayloadsForPdfGeneration_(normalized, monthCache);
  const aggregateApplied = applyAggregateInvoiceRulesFromBankFlags_(normalized, monthCache);
  const targetPatientIds = normalizeInvoicePatientIdsForGeneration_(opts.invoicePatientIds);
  const receiptEnriched = attachPreviousReceiptAmounts_(aggregateApplied, monthCache, {
    targetPatientIds
  });
  if (!receiptEnriched || !receiptEnriched.billingJson) {
    throw new Error('請求集計結果が見つかりません。先に集計を実行してください。');
  }
  const targetBillingRows = filterBillingJsonForInvoice_(
    (receiptEnriched.billingJson || []).filter(row => !(row && row.skipInvoice)),
    targetPatientIds
  );
  const invoiceTargets = includeInsurancePdf
    ? targetBillingRows.filter(row => shouldGenerateInvoicePdfForEntry_(row, receiptEnriched))
    : [];
  const receiptSummaryMap = includeInsurancePdf
    ? buildReceiptSummaryMap_(receiptEnriched, monthCache, {
      targetPatientIds
    })
    : null;
  const invoiceContexts = includeInsurancePdf
    ? invoiceTargets.map(row => {
      const pid = billingNormalizePatientId_(row && row.patientId);
      const receiptEntry = pid
        ? (receiptEnriched.billingJson || []).find(item => billingNormalizePatientId_(item && item.patientId) === pid) || row
        : row;
      const insuranceEntry = buildInsuranceInvoiceEntryForPdf_(receiptEntry) || receiptEntry;
      return buildInvoicePdfContextForEntry_(insuranceEntry, receiptEnriched, monthCache, receiptSummaryMap);
    }).filter(Boolean)
    : [];
  // 'scheduled' represents bank-flag-driven aggregate entries that skipped standard
  // invoices earlier and are now eligible for aggregate PDFs without additional triggers.
  const aggregateTargets = (receiptEnriched.billingJson || []).filter(
    row => row
      && row.skipInvoice
      && row.bankFlags
      && row.bankFlags.af === true
  );
  const aggregateFileOptions = Object.assign({}, opts, { billingMonth: normalized.billingMonth });
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
    const aggregateTotals = resolveAggregateEntryTotalsForMonths_(uniqueAggregateMonths, pid, baseEntry, monthCache);
    const aggregateRemark = formatAggregateBillingRemark_(uniqueAggregateMonths);
    const aggregateEntry = Object.assign({}, baseEntry, {
      billingMonth: normalized.billingMonth,
      receiptMonths: uniqueAggregateMonths,
      aggregateRemark,
      visitCount: aggregateTotals.visitCount,
      treatmentAmount: aggregateTotals.treatmentAmount,
      transportAmount: aggregateTotals.transportAmount,
      billingAmount: aggregateTotals.billingAmount,
      total: aggregateTotals.total,
      grandTotal: aggregateTotals.grandTotal,
      aggregateTargetMonths: uniqueAggregateMonths
    });
    const aggregateContext = buildAggregateInvoicePdfContext_(aggregateEntry, uniqueAggregateMonths, receiptEnriched, monthCache);
    const meta = generateAggregateInvoicePdf(aggregateContext, aggregateFileOptions);
    return Object.assign({}, meta, { patientId: aggregateEntry.patientId, nameKanji: aggregateEntry.nameKanji });
  }).filter(Boolean);
  const matchedIds = new Set(
    targetBillingRows
      .concat(aggregateFiles.map(file => ({ patientId: file && file.patientId })))
      .map(row => String(row && row.patientId ? row.patientId : '').trim())
      .filter(Boolean)
  );
  const missingPatientIds = targetPatientIds.filter(id => !matchedIds.has(id));

  const outputOptions = Object.assign({}, opts, { billingMonth: normalized.billingMonth, patientIds: targetPatientIds });
  const pdfs = includeInsurancePdf
    ? generateInvoicePdfs(invoiceContexts, outputOptions)
    : { files: [] };
  const insuranceFileNamesByPatient = includeInsurancePdf
    ? pdfs.files.reduce((map, meta) => {
      const pid = String(meta && meta.patientId ? meta.patientId : '').trim();
      if (pid && meta && meta.name) {
        map[pid] = meta.name;
      }
      return map;
    }, {})
    : {};
  const selfPayTargets = includeSelfPayPdf
    ? targetBillingRows.filter(row => shouldGenerateSelfPayInvoicePdfForEntry_(row))
    : [];
  const selfPayFileOptions = includeSelfPayPdf
    ? Object.assign({}, outputOptions, {
      clearMonthFolder: false
    })
    : null;
  const selfPayFiles = includeSelfPayPdf
    ? selfPayTargets.map(row => {
      const derivedEntry = buildSelfPayInvoiceEntryForPdf_(row);
      const context = buildSelfPayInvoicePdfContextForEntry_(derivedEntry, receiptEnriched, monthCache);
      if (!context) return null;
      const pid = String(context.patientId || '').trim();
      const baseFileName = insuranceFileNamesByPatient[pid] || formatInvoiceFileName_(context, {});
      const fileName = appendSelfPaySuffixToFileName_(baseFileName);
      const meta = generateInvoicePdf(context, Object.assign({}, selfPayFileOptions, { fileName }));
      return Object.assign({}, meta, { patientId: context.patientId, nameKanji: context.name });
    }).filter(Boolean)
    : [];
  const bankOutput = null;
  return {
    billingMonth: normalized.billingMonth,
    billingJson: receiptEnriched.billingJson,
    receiptStatus: normalized.receiptStatus || '',
    aggregateUntilMonth: normalized.aggregateUntilMonth || '',
    files: pdfs.files.concat(aggregateFiles, selfPayFiles),
    invoicePatientIds: targetPatientIds,
    missingInvoicePatientIds: missingPatientIds,
    invoiceGenerationMode: targetPatientIds.length ? 'partial' : 'bulk',
    bankOutput,
    preparedAt: normalized.preparedAt || null
  };
}

function generateInvoicesFromCache(billingMonth, options) {
  const month = normalizeBillingMonthInput(billingMonth);
  const loaded = loadPreparedBillingForPdfGeneration_(month.key, { withValidation: true });
  const validation = loaded && loaded.validation ? loaded.validation : null;
  const prepared = normalizePreparedBilling_(loaded && loaded.prepared);
  if (!prepared || !prepared.billingJson || (validation && validation.ok === false)) {
    throw new Error('事前集計が見つかりません。先に「請求データを集計」ボタンを実行してください。');
  }
  const monthCache = createBillingMonthCache_();
  loadPreparedBillingSummariesIntoCache_(monthCache);
  loadBankWithdrawalAmountsIntoCache_(monthCache, prepared);
  return generatePreparedInvoices_(prepared, Object.assign({}, options || {}, { monthCache }));
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

  const loaded = loadPreparedBillingForPdfGeneration_(month.key, { withValidation: true });
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
  if (monthCache.preparedByMonth) {
    monthCache.preparedByMonth[month.key] = reducePreparedBillingSummary_(prepared);
  }
  loadPreparedBillingSummariesIntoCache_(monthCache);
  loadBankWithdrawalAmountsIntoCache_(monthCache, prepared);
  preloadPreparedPayloadsForPdfGeneration_(prepared, monthCache);
  const receiptPrepared = attachPreviousReceiptAmounts_(prepared, monthCache, {
    targetPatientIds: [patientId]
  });
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
    const hasManualBillingInput = edit && (Object.prototype.hasOwnProperty.call(edit, 'manualBillingAmount')
      || Object.prototype.hasOwnProperty.call(edit, 'grandTotal')
      || Object.prototype.hasOwnProperty.call(edit, 'billingAmount'));
    const manualBillingSource = hasManualBillingInput
      ? (Object.prototype.hasOwnProperty.call(edit, 'manualBillingAmount')
        ? edit.manualBillingAmount
        : (Object.prototype.hasOwnProperty.call(edit, 'grandTotal') ? edit.grandTotal : edit.billingAmount))
      : undefined;
    const normalizedManualBilling = manualBillingSource === '' || manualBillingSource === null
      ? ''
      : Number(manualBillingSource) || 0;
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
      manualBillingAmount: hasManualBillingInput ? normalizedManualBilling : undefined,
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
    onlineConsent: edit.onlineConsent,
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
    manualSelfPayAmount: edit.manualSelfPayAmount,
    manualBillingAmount: edit.manualBillingAmount
  };
}

function hasBillingOverrideValues_(edit) {
  const fields = extractBillingOverrideFields_(edit);
  return ['manualUnitPrice', 'manualTransportAmount', 'carryOverAmount', 'adjustedVisitCount', 'manualSelfPayAmount', 'manualBillingAmount']
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
  const colOnlineConsent = resolveBillingColumn_(headers, ['オンライン同意', 'オンライン同意フラグ', 'オンライン同意（AT）', 'オンライン同意(AT)'], 'オンライン同意', { fallbackLetter: 'AT' });
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
    if (colOnlineConsent && fields.onlineConsent !== undefined) newRow[colOnlineConsent - 1] = normalizeZeroOneFlag_(fields.onlineConsent);

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
  if (sheet) {
    const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
    if (lastCol > 0) {
      const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
      if (headers.indexOf('manualBillingAmount') === -1) {
        sheet.getRange(1, lastCol + 1).setValue('manualBillingAmount');
      }
      if (headers.indexOf('entryType') === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue('entryType');
      }
    }
    return sheet;
  }

  sheet = ss.insertSheet(BILLING_OVERRIDES_SHEET_NAME);
  sheet.appendRow([
    'ym',
    'patientId',
    'entryType',
    'manualUnitPrice',
    'manualTransportAmount',
    'carryOverAmount',
    'adjustedVisitCount',
    'manualSelfPayAmount',
    'manualBillingAmount'
  ]);
  return sheet;
}

function resolveBillingOverridesColumns_(headers) {
  const colYm = resolveBillingColumn_(headers, ['ym', 'billingMonth', 'month'], 'ym', { required: true });
  const colPid = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', 'patientId']), 'patientId', { required: true });
  const colEntryType = resolveBillingColumn_(headers, ['entryType', 'entry_type', 'type', '請求タイプ', '請求種別'], 'entryType', {});
  const colManualUnitPrice = resolveBillingColumn_(headers, ['manualUnitPrice', 'unitPrice', '単価'], 'manualUnitPrice', {});
  const colManualTransport = resolveBillingColumn_(headers, ['manualTransportAmount', 'transportAmount', '交通費'], 'manualTransportAmount', {});
  const colCarryOver = resolveBillingColumn_(headers, ['carryOverAmount', 'carryOver', '未入金額', '繰越'], 'carryOverAmount', {});
  const colAdjustedVisits = resolveBillingColumn_(headers, ['adjustedVisitCount', 'visitCount', '回数'], 'adjustedVisitCount', {});
  const colManualSelfPay = resolveBillingColumn_(headers, ['manualSelfPayAmount', 'selfPayAmount', '自費'], 'manualSelfPayAmount', {});
  const colManualBillingAmount = resolveBillingColumn_(
    headers,
    ['manualBillingAmount', 'billingAmount', 'grandTotal', '請求額', '合計'],
    'manualBillingAmount',
    {}
  );
  return {
    colYm,
    colPid,
    colEntryType,
    colManualUnitPrice,
    colManualTransport,
    colCarryOver,
    colAdjustedVisits,
    colManualSelfPay,
    colManualBillingAmount
  };
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
    assignField(newRow, cols.colEntryType, edit.entryType);
    assignField(newRow, cols.colManualUnitPrice, edit.manualUnitPrice);
    assignField(newRow, cols.colManualTransport, edit.manualTransportAmount);
    assignField(newRow, cols.colCarryOver, edit.carryOverAmount);
    assignField(newRow, cols.colAdjustedVisits, edit.adjustedVisitCount);
    assignField(newRow, cols.colManualSelfPay, edit.manualSelfPayAmount);
    assignField(newRow, cols.colManualBillingAmount, edit.manualBillingAmount);

    sheet.getRange(idx + 2, 1, 1, newRow.length).setValues([newRow]);
    return { updated: true, rowNumber: idx + 2 };
  }

  const newRow = new Array(Math.max(lastCol, 9)).fill('');
  assignField(newRow, cols.colYm, ym);
  assignField(newRow, cols.colPid, pid);
  assignField(newRow, cols.colEntryType, edit.entryType);
  assignField(newRow, cols.colManualUnitPrice, edit.manualUnitPrice);
  assignField(newRow, cols.colManualTransport, edit.manualTransportAmount);
  assignField(newRow, cols.colCarryOver, edit.carryOverAmount);
  assignField(newRow, cols.colAdjustedVisits, edit.adjustedVisitCount);
  assignField(newRow, cols.colManualSelfPay, edit.manualSelfPayAmount);
  assignField(newRow, cols.colManualBillingAmount, edit.manualBillingAmount);

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
    const normalizedManualBilling = edit.manualBillingAmount === '' || edit.manualBillingAmount === null
      ? ''
      : edit.manualBillingAmount;
    const result = saveBillingOverrideUpdate_(billingMonthKey, Object.assign({}, edit, {
      carryOverAmount: normalizedCarryOver,
      manualUnitPrice: normalizedManualUnitPrice,
      manualTransportAmount: normalizedManualTransport,
      manualSelfPayAmount: normalizedManualSelfPay,
      manualBillingAmount: normalizedManualBilling
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
  const normalizedPrepared = normalizePreparedBilling_(prepared) || prepared;
  savePreparedBilling_(normalizedPrepared);
  savePreparedBillingToSheet_(billingMonth, normalizedPrepared);
  return normalizedPrepared;
}

function applyBillingEditsAndGenerateInvoices(billingMonth, options) {
  const prepared = applyBillingEdits(billingMonth, options);
  const monthCache = createBillingMonthCache_();
  loadPreparedBillingSummariesIntoCache_(monthCache);
  loadBankWithdrawalAmountsIntoCache_(monthCache, prepared);
  return generatePreparedInvoices_(prepared, Object.assign({}, options || {}, { monthCache }));
}

function generatePreparedInvoicesForMonth(billingMonth, options) {
  const normalizedMonth = normalizeBillingMonthInput(billingMonth);
  const monthContext = normalizedMonth
    ? { key: normalizedMonth.key, year: normalizedMonth.year, month: normalizedMonth.month }
    : null;
  const monthKey = monthContext && monthContext.key ? monthContext.key : '';
  if (!monthKey) {
    throw new Error('PDF対象月が指定されていません。');
  }

  const opts = options || {};
  if (opts.applyEdits === true) {
    billingLogger_.log('[billing] generatePreparedInvoicesForMonth ignored applyEdits to keep PDF generation read-only');
  }

  const loaded = loadPreparedBillingForPdfGeneration_(monthKey, { withValidation: true });
  const validation = loaded && loaded.validation ? loaded.validation : null;
  const prepared = normalizePreparedBilling_(loaded && loaded.prepared);
  if (!prepared || !prepared.billingJson || (validation && validation.ok === false)) {
    throw new Error('事前集計が見つかりません。先に「請求データを集計」ボタンを実行してください。');
  }

  const monthCache = createBillingMonthCache_();
  if (monthCache.preparedByMonth) {
    monthCache.preparedByMonth[monthKey] = reducePreparedBillingSummary_(prepared);
  }
  loadPreparedBillingSummariesIntoCache_(monthCache);
  loadBankWithdrawalAmountsIntoCache_(monthCache, prepared);
  preloadPreparedPayloadsForPdfGeneration_(prepared, monthCache);

  return generatePreparedInvoices_(prepared, Object.assign({}, opts, { monthCache }));
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
    const normalized = normalizePreparedBilling_(prepared) || prepared;
    savePreparedBilling_(normalized);
    savePreparedBillingToSheet_(billingMonth, normalized);
    return exportBankTransferDataForPrepared_(normalized, options || {});
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
