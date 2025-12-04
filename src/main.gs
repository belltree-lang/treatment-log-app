/**
 * Billing entry points for Apps Script deployment.
 *
 * These helper functions wire the Get/Logic/Output layers together so that
 * callers can preview JSON or generate patient invoice PDFs for a given billing month.
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('billing');
  template.baseUrl = ScriptApp.getService().getUrl() || '';
  return template
    .evaluate()
    .setTitle('請求処理アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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

const BILLING_CACHE_PREFIX = 'billing_prepared_';
const BILLING_CACHE_TTL_SECONDS = 3600; // 1 hour

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

function validatePreparedBillingPayload_(payload, expectedMonthKey) {
  if (!payload || typeof payload !== 'object') return { ok: false, reason: 'payload missing' };
  const billingMonth = payload.billingMonth || payload.month || '';
  if (!billingMonth) return { ok: false, reason: 'billingMonth missing' };
  if (expectedMonthKey && String(billingMonth) !== String(expectedMonthKey)) {
    return { ok: false, reason: 'billingMonth mismatch' };
  }
  if (!Array.isArray(payload.billingJson)) return { ok: false, reason: 'billingJson missing' };
  return { ok: true, billingMonth };
}

function loadPreparedBilling_(billingMonthKey) {
  const key = buildBillingCacheKey_(billingMonthKey);
  if (!key) return null;
  const cache = getBillingCache_();
  if (!cache) return null;
  const cached = cache.get(key);
  if (!cached) {
    try {
      Logger.log('[billing] loadPreparedBilling_: cache miss for ' + key);
    } catch (err) {
      // ignore logging errors in non-GAS environments
    }
    return null;
  }
  try {
    Logger.log('[billing] loadPreparedBilling_ raw cache for ' + key + ': ' + cached);
  } catch (err) {
    // ignore logging errors in non-GAS environments
  }
  try {
    const parsed = JSON.parse(cached);
    const validation = validatePreparedBillingPayload_(parsed, billingMonthKey);
    if (!validation.ok) {
      try {
        Logger.log('[billing] loadPreparedBilling_: invalid cache for ' + key + ' reason=' + validation.reason);
      } catch (err) {
        // ignore logging errors in non-GAS environments
      }
      console.warn('[billing] Prepared cache invalid for ' + key + ': ' + validation.reason);
      clearBillingCache_(key);
      return null;
    }
    try {
      Logger.log('[billing] loadPreparedBilling_: parsed cache billingJson length=' + (parsed.billingJson || []).length);
    } catch (err) {
      // ignore logging errors in non-GAS environments
    }
    return Object.assign({}, parsed, { billingMonth: validation.billingMonth });
  } catch (err) {
    try {
      Logger.log('[billing] loadPreparedBilling_: failed to parse cache for ' + key + ' error=' + err);
    } catch (logErr) {
      // ignore logging errors in non-GAS environments
    }
    console.warn('[billing] Failed to parse prepared cache', err);
    clearBillingCache_(key);
    return null;
  }
}

function savePreparedBilling_(payload) {
  const key = buildBillingCacheKey_(payload && payload.billingMonth);
  if (!key) return;
  const cache = getBillingCache_();
  if (!cache) return;
  try {
    cache.put(key, JSON.stringify(payload), BILLING_CACHE_TTL_SECONDS);
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
  const source = getBillingSourceData(billingMonth);
  const billingJson = generateBillingJsonFromSource(source);
  const visitCounts = source.treatmentVisitCounts || source.visitCounts || {};
  const visitCountKeys = Object.keys(visitCounts || {});
  const zeroVisitSamples = visitCountKeys
    .filter(pid => {
      const entry = visitCounts[pid];
      const visitCount = entry && entry.visitCount != null ? entry.visitCount : entry;
      return !visitCount || Number(visitCount) === 0;
    })
    .slice(0, 5);

  billingLogger_.log('[billing] buildPreparedBillingPayload_ summary=' + JSON.stringify({
    billingMonth: source.billingMonth,
    treatmentVisitCountEntries: visitCountKeys.length,
    zeroVisitSamples,
    billingJsonLength: Array.isArray(billingJson) ? billingJson.length : null
  }));
  if (Array.isArray(billingJson) && billingJson.length) {
    billingLogger_.log('[billing] buildPreparedBillingPayload_ firstBillingEntry=' + JSON.stringify(billingJson[0]));
  }
  return {
    billingMonth: source.billingMonth,
    billingJson,
    preparedAt: new Date().toISOString(),
    patients: source.patients || source.patientMap || {},
    bankInfoByName: source.bankInfoByName || {},
    staffByPatient: source.staffByPatient || {},
    staffDirectory: source.staffDirectory || {},
    staffDisplayByPatient: source.staffDisplayByPatient || {}
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
  if (!payload) return null;
  return Object.assign({}, payload, { billingJson: coerceBillingJsonArray_(payload.billingJson) });
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
    billingMonth: normalized.billingMonth || '',
    billingJson,
    preparedAt: normalized.preparedAt || null,
    patients: normalized.patients || {},
    bankInfoByName: normalized.bankInfoByName || {},
    staffByPatient: normalized.staffByPatient || {},
    staffDirectory: normalized.staffDirectory || {},
    staffDisplayByPatient: normalized.staffDisplayByPatient || {}
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
  const prepared = buildPreparedBillingPayload_(billingMonth);
  const clientPayload = toClientBillingPayload_(prepared);
  const serialized = serializeBillingPayload_(clientPayload) || clientPayload;
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

  savePreparedBilling_(serialized);
  return serialized;
}

/**
 * Generate billing JSON without creating files (preview use case).
 * @param {string|Date|Object} billingMonth - YYYYMM string, Date, or normalized month object.
 * @return {Object} billingMonth key and generated JSON array.
 */
function generateBillingJsonPreview(billingMonth) {
  const prepared = buildPreparedBillingPayload_(billingMonth);
  savePreparedBilling_(prepared);
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
    const hasManualUnitPriceInput = edit && Object.prototype.hasOwnProperty.call(edit, 'unitPrice');
    const hasManualUnitPrice = hasManualUnitPriceInput && edit.unitPrice !== null && edit.unitPrice !== '';
    const hasManualTransportInput = edit && (Object.prototype.hasOwnProperty.call(edit, 'manualTransportAmount')
      || Object.prototype.hasOwnProperty.call(edit, 'transportAmount'));
    const manualTransportSource = edit && Object.prototype.hasOwnProperty.call(edit, 'manualTransportAmount')
      ? edit.manualTransportAmount
      : edit.transportAmount;
    const normalizedManualTransport = manualTransportSource === '' || manualTransportSource === null
      ? ''
      : Number(manualTransportSource) || 0;
    return {
      patientId: pid,
      insuranceType: edit.insuranceType != null ? String(edit.insuranceType).trim() : undefined,
      medicalAssistance: normalizeBillingEditMedicalAssistance_(edit.medicalAssistance),
      burdenRate: burden !== null ? burden : undefined,
      // normalized unitPrice retains legacy behavior for immediate billing recalculation,
      // while manualUnitPrice preserves blank input for persistence.
      unitPrice: hasManualUnitPrice ? Number(edit.unitPrice) || 0 : undefined,
      manualUnitPrice: hasManualUnitPriceInput ? edit.unitPrice : undefined,
      carryOverAmount: edit.carryOverAmount != null ? Number(edit.carryOverAmount) || 0 : undefined,
      payerType: edit.payerType != null ? String(edit.payerType).trim() : undefined,
      manualTransportAmount: hasManualTransportInput ? normalizedManualTransport : undefined,
      responsible: edit && edit.responsible != null ? String(edit.responsible).trim() : undefined,
      bankCode: normalizedBankCode,
      branchCode: normalizedBranchCode,
      accountNumber: normalizedAccountNumber,
      isNew: edit && Object.prototype.hasOwnProperty.call(edit, 'isNew')
        ? normalizeZeroOneFlag_(edit.isNew)
        : undefined
    };
  }).filter(Boolean);
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
  const colMedical = resolveBillingColumn_(headers, ['医療助成'], '医療助成', { fallbackLetter: 'AS' });
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
    if (colMedical && fields.medicalAssistance !== undefined) newRow[colMedical - 1] = normalizeBillingEditMedicalAssistance_(fields.medicalAssistance) ? 1 : 0;
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
    const result = savePatientUpdate(edit.patientId, {
      insuranceType: edit.insuranceType,
      burdenRate: edit.burdenRate,
      medicalAssistance: edit.medicalAssistance,
      manualUnitPrice: edit.manualUnitPrice,
      manualTransportAmount: edit.manualTransportAmount,
      carryOverAmount: edit.carryOverAmount,
      payerType: edit.payerType,
      responsible: edit.responsible,
      bankCode: edit.bankCode,
      branchCode: edit.branchCode,
      accountNumber: edit.accountNumber,
      isNew: edit.isNew
    });
    if (result && result.updated) {
      updatedCount += 1;
    }
  });

  return { updated: updatedCount };
}

function applyBillingEdits(billingMonth, options) {
  const opts = options || {};
  const edits = normalizeBillingEdits_(opts.edits);
  let refreshedPatients = null;
  if (edits.length) {
    applyBillingPatientEdits_(edits);
    refreshedPatients = getBillingPatientRecords();
  }
  const prepared = buildPreparedBillingPayload_(billingMonth);
  if (refreshedPatients) {
    prepared.patients = indexByPatientId_(refreshedPatients);
  }
  savePreparedBilling_(prepared);
  return prepared;
}

function applyBillingEditsAndGenerateInvoices(billingMonth, options) {
  const prepared = applyBillingEdits(billingMonth, options);
  return generatePreparedInvoices_(prepared, options || {});
}

function generateBankTransferData(billingMonth, options) {
  const prepared = buildPreparedBillingPayload_(billingMonth);
  savePreparedBilling_(prepared);
  return exportBankTransferDataForPrepared_(prepared, options || {});
}

function generateBankTransferDataFromCache(billingMonth, options) {
  const month = normalizeBillingMonthInput(billingMonth);
  const prepared = loadPreparedBilling_(month.key);
  if (!prepared || !prepared.billingJson) {
    throw new Error('事前集計が見つかりません。先に「請求データを集計」を実行してください。');
  }
  return exportBankTransferDataForPrepared_(prepared, options || {});
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
