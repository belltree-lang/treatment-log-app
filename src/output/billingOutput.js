/***** Output layer: billing invoice PDF generation *****/

const INVOICE_PARENT_FOLDER_ID = '1EG-GB3PbaUr9C1LJWlaf_idqoYF-19Ux';
const INVOICE_FILE_PREFIX = '請求書';
const TRANSPORT_PRICE = (typeof BILLING_TRANSPORT_UNIT_PRICE !== 'undefined')
  ? BILLING_TRANSPORT_UNIT_PRICE
  : 33;
const INVOICE_TREATMENT_UNIT_PRICE_BY_BURDEN = { 1: 417, 2: 834, 3: 1251 };
const INVOICE_UNIT_PRICE_FALLBACK = (typeof BILLING_UNIT_PRICE !== 'undefined') ? BILLING_UNIT_PRICE : 4170;

function escapeHtml_(value) {
  return String(value || '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[c]);
}

const normalizeInvoiceBurdenRateInt_ = typeof normalizeBurdenRateInt_ === 'function'
  ? normalizeBurdenRateInt_
  : function fallbackNormalizeInvoiceBurdenRateInt_(burdenRate) {
    if (burdenRate == null || burdenRate === '') return 0;
    if (String(burdenRate).trim() === '自費') return '自費';
    const num = Number(burdenRate);
    if (Number.isFinite(num)) {
      if (num > 0 && num < 1) return Math.round(num * 10);
    if (num >= 1 && num < 10) return Math.round(num);
    if (num >= 10 && num <= 100) return Math.round(num / 10);
  }
  const normalized = String(burdenRate).normalize('NFKC').replace(/\s+/g, '').replace('％', '%');
  const hasPercent = normalized.indexOf('%') >= 0;
  const numericText = normalized.replace(/[^0-9.]/g, '');
  const parsed = Number(numericText);
  if (!Number.isFinite(parsed)) return 0;
  if (parsed === 0) return 0;
  if (hasPercent) return Math.round(parsed / 10);
  if (parsed > 0 && parsed < 10) return Math.round(parsed);
  if (parsed >= 10 && parsed <= 100) return Math.round(parsed / 10);
  return 0;
  };

const resolveInvoiceUnitPriceForOutput_ = typeof resolveInvoiceUnitPrice_ === 'function'
  ? resolveInvoiceUnitPrice_
  : function fallbackResolveInvoiceUnitPriceForOutput_(insuranceType, burdenRate, manualUnitPrice, medicalAssistance, patientUnitPrice) {
    const type = String(insuranceType || '').trim();
    if (type === 'マッサージ') return 0;
    const normalizedManual = normalizeInvoiceMoney_(manualUnitPrice);
    const hasManual = Number.isFinite(normalizedManual) && normalizedManual !== 0;
    if (hasManual) return normalizedManual;
    const assistance = normalizeInvoiceMedicalAssistanceFlag_(medicalAssistance);
    if (assistance === 1) return 0;
    const isLifeProtection = ['生保', '生活保護', '生活扶助'].indexOf(type) >= 0;
    if (isLifeProtection) return 0;
    const normalizedBurdenRate = normalizeInvoiceBurdenRateInt_(burdenRate);
    const isSelfPaid = type === '自費' || normalizedBurdenRate === '自費';
    if (isSelfPaid) return 0;
    const normalizedPatientPrice = normalizeInvoiceMoney_(patientUnitPrice);
    if (Number.isFinite(normalizedPatientPrice) && normalizedPatientPrice !== 0) return normalizedPatientPrice;
    return INVOICE_UNIT_PRICE_FALLBACK;
  };

function convertSpreadsheetToExcelBlob_(file, exportName) {
  if (!file || typeof file.getMimeType !== 'function') {
    throw new Error('スプレッドシート以外のファイルをExcelに変換することはできません');
  }

  const mimeType = file.getMimeType();
  const isSpreadsheet = mimeType === MimeType.GOOGLE_SHEETS;
  const isExcel = mimeType === MimeType.MICROSOFT_EXCEL;
  if (!isSpreadsheet && !isExcel) {
    throw new Error('スプレッドシート以外のファイルをExcelに変換することはできません');
  }

  const blob = file.getBlob();
  const name = (exportName && String(exportName).trim()) || 'export';
  const excelBlob = isSpreadsheet && typeof blob.getAs === 'function'
    ? blob.getAs(MimeType.MICROSOFT_EXCEL)
    : blob;
  return excelBlob.setName(name + '.xlsx');
}

function resolveBillingSpreadsheetForOutput_() {
  const root = typeof globalThis !== 'undefined' ? globalThis : this;

  if (root && typeof root.billingSs === 'function') {
    try {
      return root.billingSs();
    } catch (err) {
      return null;
    }
  }

  if (root && typeof root.ss === 'function') {
    try {
      return root.ss();
    } catch (err) {
      return null;
    }
  }

  return null;
}

function normalizeBillingAmount_(item) {
  if (!item) return 0;

  if (item.grandTotal != null && item.grandTotal !== '') {
    return normalizeInvoiceMoney_(item.grandTotal);
  }

  const carryOverTotal = normalizeInvoiceMoney_(item.carryOverAmount)
    + normalizeInvoiceMoney_(item.carryOverFromHistory);

  if (item.total != null && item.total !== '') {
    return normalizeInvoiceMoney_(item.total) + carryOverTotal;
  }

  const billingAmount = normalizeInvoiceMoney_(item.billingAmount);
  const treatmentAmount = normalizeInvoiceMoney_(item.treatmentAmount);
  const transportAmount = normalizeInvoiceMoney_(item.transportAmount);

  if (item.billingAmount != null && item.billingAmount !== '') {
    return billingAmount + transportAmount + carryOverTotal;
  }

  if (item.treatmentAmount != null || item.transportAmount != null || carryOverTotal) {
    return treatmentAmount + transportAmount + carryOverTotal;
  }

  return 0;
}

const resolveBillingAmountForMonthAndPatientForOutput_ = (typeof globalThis !== 'undefined'
  && typeof globalThis.resolveBillingAmountForMonthAndPatient_ === 'function')
  ? globalThis.resolveBillingAmountForMonthAndPatient_
  : function fallbackResolveBillingAmountForMonthAndPatientForOutput_(billingMonth, patientId, fallbackEntry) {
    const normalizedMonth = normalizeInvoiceMonthKey_(billingMonth);
    const fallbackMonth = fallbackEntry && normalizeInvoiceMonthKey_(fallbackEntry.billingMonth);
    if (normalizedMonth && fallbackMonth && normalizedMonth === fallbackMonth) {
      return normalizeBillingAmount_(fallbackEntry);
    }
    return 0;
  };

function normalizeBillingNameKey_(value) {
  return String(value || '')
    .normalize('NFKC')
    .replace(/[\s\u3000・･．.ー－ｰ−‐-]+/g, '')
    .trim();
}

function normalizeBillingFullNameKey_(nameKanji, nameKana) {
  const kanjiKey = normalizeBillingNameKey_(nameKanji);
  const kanaKey = normalizeBillingNameKey_(nameKana);
  const combined = [kanjiKey, kanaKey].filter(Boolean).join('::');
  if (!combined) return '';
  const numericOnly = combined.replace(/::/g, '');
  if (/^\d+$/.test(numericOnly)) return '';
  return combined;
}

/**
 * 氏名に紐づく統一キーを生成する。
 * - 正式キー: 漢字・カナを NFKC 正規化 + 空白除去し `kanji::kana` 形式で連結
 * - フォールバック: どちらか一方のみの氏名がある場合はその値を同じ手順で正規化
 * - 最終手段: patientId を空白除去したものをキーとして返す
 */
function buildBillingNameKey_(record) {
  if (!record) return '';
  const fullNameKey = normalizeBillingFullNameKey_(record.nameKanji, record.nameKana);
  if (fullNameKey) return fullNameKey;

  const singleNameKey = normalizeBillingNameKey_(record.nameKanji || record.nameKana);
  if (singleNameKey) return singleNameKey;

  const storedKey = normalizeBillingNameKey_(record._nameKey);
  if (storedKey) return storedKey;

  if (record.patientId != null) {
    const patientIdKey = normalizeBillingNameKey_(record.patientId);
    if (patientIdKey) return patientIdKey;
  }

  return '';
}

function normalizeKana_(value) {
  return String(value || '')
    .normalize('NFKC')
    .trim();
}

function normalizeInvoiceMoney_(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }
  const text = String(value || '')
    .normalize('NFKC')
    .replace(/[，,]/g, '')
    .trim();
  if (!text) return 0;
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
}

function normalizeInvoiceVisitCount_(value) {
  const source = value && value.visitCount != null ? value.visitCount : value;
  if (typeof source === 'number') {
    return Number.isFinite(source) && source > 0 ? source : 0;
  }
  const normalized = String(source || '')
    .normalize('NFKC')
    .replace(/[，,]/g, '')
    .trim();
  const num = Number(normalized);
  return Number.isFinite(num) && num > 0 ? num : 0;
}

function normalizeInvoiceMedicalAssistanceFlag_(value) {
  if (value === 1 || value === '1' || value === true) return 1;
  return 0;
}

function normalizeBillingCarryOver_(item) {
  if (!item) return 0;
  const directCarryOver = (item.carryOverAmount != null && item.carryOverAmount !== '')
    ? normalizeInvoiceMoney_(item.carryOverAmount)
    : (item.raw && item.raw.carryOverAmount != null)
      ? normalizeInvoiceMoney_(item.raw.carryOverAmount)
      : 0;
  const historyCarryOver = normalizeInvoiceMoney_(item.carryOverFromHistory);
  return directCarryOver + historyCarryOver;
}

function formatBillingCurrency_(value) {
  const num = Number(value);
  if (!isFinite(num)) return '0';
  return Math.round(num).toLocaleString('ja-JP');
}

function normalizeBillingMonthLabel_(billingMonth) {
  const digits = (billingMonth ? String(billingMonth) : '').replace(/\D/g, '');
  if (digits.length >= 6) {
    const year = digits.slice(0, 4);
    const month = digits.slice(4, 6).padStart(2, '0');
    return year + '年' + month + '月';
  }
  return billingMonth || '';
}

function formatBillingMonthForFile_(billingMonth) {
  const digits = (billingMonth ? String(billingMonth) : '').replace(/\D/g, '');
  if (digits.length >= 6) {
    const year = digits.slice(0, 4);
    const month = digits.slice(4, 6).padStart(2, '0');
    return year + '-' + month;
  }
  return billingMonth || '';
}

function formatBillingMonthCompact_(billingMonth) {
  const digits = (billingMonth ? String(billingMonth) : '').replace(/\D/g, '');
  if (digits.length >= 6) {
    return digits.slice(0, 6);
  }
  return '';
}

function normalizeInvoiceMonthKey_(value) {
  const digits = (value ? String(value) : '').replace(/\D/g, '');
  if (digits.length >= 6) return digits.slice(0, 6);
  return '';
}

function formatMonthWithReiwaEra_(yyyymm) {
  const normalized = normalizeInvoiceMonthKey_(yyyymm);
  if (!normalized) return '';
  const year = Number(normalized.slice(0, 4));
  const month = normalized.slice(4, 6);
  const eraYear = Number.isFinite(year) ? year - 2018 : '';
  if (!eraYear) return '';
  return `令和${eraYear}年${month}月`;
}

function buildInvoiceChargePeriodLabel_(data) {
  const months = [];
  const pushMonth = (value) => {
    const key = normalizeInvoiceMonthKey_(value);
    if (key) months.push(key);
  };

  const amount = data && data.amount ? data.amount : data;
  const isAggregate = data && (data.isAggregateInvoice || data.invoiceMode === 'aggregate');
  if (isAggregate && Array.isArray(amount && amount.aggregateMonthTotals)) {
    amount.aggregateMonthTotals.forEach(row => pushMonth(row && row.month));
  }

  if (isAggregate && !months.length && Array.isArray(amount && amount.receiptMonths)) {
    amount.receiptMonths.forEach(pushMonth);
  }

  if (isAggregate && !months.length && Array.isArray(data && data.months)) {
    data.months.forEach(pushMonth);
  }

  if (!months.length && data && data.billingMonth) {
    pushMonth(data.billingMonth);
  }

  const uniqueSorted = Array.from(new Set(months)).sort();
  if (!uniqueSorted.length) return '';

  const start = uniqueSorted[0];
  const end = uniqueSorted[uniqueSorted.length - 1];
  const startLabel = formatMonthWithReiwaEra_(start);
  if (!startLabel) return '';
  if (start === end) return `${startLabel}分`;

  const startYear = start.slice(0, 4);
  const endYear = end.slice(0, 4);
  const endLabel = startYear === endYear
    ? (end.slice(4, 6) ? `${end.slice(4, 6)}月` : '')
    : formatMonthWithReiwaEra_(end);
  if (!endLabel) return `${startLabel}分`;

  return `${startLabel}分〜${endLabel}分`;
}

function buildInclusiveMonthRange_(fromYm, toYm) {
  const startKey = normalizeInvoiceMonthKey_(fromYm);
  const endKey = normalizeInvoiceMonthKey_(toYm);
  if (!startKey || !endKey) return startKey ? [startKey] : [];
  const months = [];
  const startNum = Number(startKey);
  const endNum = Number(endKey);
  if (!Number.isFinite(startNum) || !Number.isFinite(endNum) || startNum > endNum) return [startKey];

  let cursorYear = Number(startKey.slice(0, 4));
  let cursorMonth = Number(startKey.slice(4, 6));
  while (true) {
    const ym = String(cursorYear).padStart(4, '0') + String(cursorMonth).padStart(2, '0');
    months.push(ym);
    if (ym === endKey) break;
    cursorMonth += 1;
    if (cursorMonth > 12) {
      cursorMonth = 1;
      cursorYear += 1;
    }
    if (months.length > 240) break;
  }
  return months;
}

function formatAggregatedReceiptRemark_(months) {
  if (!Array.isArray(months) || !months.length) return '';
  const parts = months.map((ym, idx) => {
    const label = formatMonthWithReiwaEra_(ym);
    if (!label) return '';
    if (idx === 0) return label + '分';
    const month = normalizeInvoiceMonthKey_(ym).slice(4, 6);
    return month ? month + '月分' : '';
  }).filter(Boolean);

  if (!parts.length) return '';
  return parts.join('・') + '施術料金として';
}

function normalizeReceiptMonths_(months) {
  const list = Array.isArray(months) ? months : [];
  const seen = new Set();
  const normalized = list.map(value => normalizeInvoiceMonthKey_(value)).filter(Boolean).filter(ym => {
    if (seen.has(ym)) return false;
    seen.add(ym);
    return true;
  });

  return normalized;
}

function normalizePastInvoiceMonths_(months, billingMonth) {
  const list = Array.isArray(months) ? months : [];
  const billingKey = normalizeInvoiceMonthKey_(billingMonth);
  const billingNum = Number(billingKey) || 0;
  const seen = new Set();
  const normalized = [];

  list.forEach(value => {
    const ym = normalizeInvoiceMonthKey_(value);
    if (!ym || seen.has(ym)) return;
    const ymNum = Number(ym) || 0;
    if (billingNum && ymNum >= billingNum) return;
    seen.add(ym);
    normalized.push(ym);
  });

  return normalized.sort();
}

function resolveHasPreviousReceiptSheet_(item) {
  if (!item) return false;
  if (Object.prototype.hasOwnProperty.call(item, 'hasPreviousReceiptSheet')) {
    return !!item.hasPreviousReceiptSheet;
  }
  if (Object.prototype.hasOwnProperty.call(item, 'hasPreviousPrepared')) {
    return !!item.hasPreviousPrepared;
  }
  return false;
}

function normalizeAggregateStatus_(status) {
  const normalized = typeof status === 'string' ? status.trim().toLowerCase() : String(status || '').trim().toLowerCase();
  return normalized || '';
}

function resolveInvoiceReceiptDisplay_(item, options) {
  const hasPreviousReceiptSheet = resolveHasPreviousReceiptSheet_(item);
  const billingMonthKey = normalizeInvoiceMonthKey_(item && item.billingMonth);
  const explicitReceiptMonths = normalizePastInvoiceMonths_(
    normalizeReceiptMonths_(item && item.receiptMonths),
    billingMonthKey
  );
  const overrideMonths = options && Array.isArray(options.aggregateMonths)
    ? options.aggregateMonths
    : [];
  const aggregateDecisionMonths = overrideMonths.length
    ? normalizeAggregateMonthsForInvoice_(overrideMonths, billingMonthKey)
    : [];
  const aggregateStatus = normalizeAggregateStatus_(item && item.aggregateStatus);
  const aggregateConfirmed = aggregateStatus === 'confirmed';
  const receiptMonths = aggregateDecisionMonths.length ? aggregateDecisionMonths : explicitReceiptMonths;
  const customReceiptRemark = item && item.receiptRemark ? String(item.receiptRemark) : '';
  const receiptRemark = customReceiptRemark || (receiptMonths.length > 1
    ? formatAggregatedReceiptRemark_(receiptMonths)
    : '');
  const receiptStatus = item && item.receiptStatus ? String(item.receiptStatus).toUpperCase() : '';
  const shouldHideByStatus = receiptStatus === 'UNPAID' || receiptStatus === 'HOLD';
  const visible = !shouldHideByStatus && !item.skipReceipt && receiptMonths.length > 0 && hasPreviousReceiptSheet;
  const receiptMonthsSource = aggregateDecisionMonths.length
    ? 'aggregateDecisionMonths'
    : (receiptMonths.length ? 'explicit' : 'none');

  return {
    visible,
    receiptRemark,
    receiptMonths,
    explicitReceiptMonths,
    aggregateDecisionMonths,
    receiptMonthsSource,
    aggregateStatus,
    aggregateConfirmed
  };
}

function resolveAggregateInvoiceDecision_(item, receipt, billingMonth, options) {
  const billingMonthKey = normalizeInvoiceMonthKey_(billingMonth);
  const explicitReceiptMonths = Array.isArray(receipt && receipt.explicitReceiptMonths)
    ? normalizeAggregateMonthsForInvoice_(receipt.explicitReceiptMonths, billingMonthKey)
    : [];
  const aggregateTargetMonths = Array.isArray(item && item.aggregateTargetMonths)
    ? normalizeAggregateMonthsForInvoice_(item.aggregateTargetMonths, billingMonthKey)
    : [];
  const requestedAggregateMonths = normalizeAggregateMonthsForInvoice_(
    explicitReceiptMonths.length ? explicitReceiptMonths : aggregateTargetMonths,
    billingMonthKey
  );
  const preparedAggregate = filterAggregateMonthsByPrepared_(
    item,
    requestedAggregateMonths,
    options
  );
  const aggregateDecisionMonths = preparedAggregate.months;
  const missingAggregateMonths = preparedAggregate.missingMonths;
  const usesExplicitReceiptMonths = explicitReceiptMonths.length > 0;
  const usesAggregateTargetMonths = !usesExplicitReceiptMonths && requestedAggregateMonths.length > 0;
  const decisionSources = [];
  if (usesExplicitReceiptMonths) decisionSources.push('explicitReceiptMonths');
  if (usesAggregateTargetMonths) decisionSources.push('aggregateTargetMonths');
  const aggregateMonthsSource = usesExplicitReceiptMonths
    ? 'explicitReceiptMonths'
    : (usesAggregateTargetMonths ? 'aggregateTargetMonths' : 'none');

  const trace = {
    billingMonth: billingMonthKey,
    receiptMonthsSource: receipt && receipt.receiptMonthsSource ? receipt.receiptMonthsSource : 'none',
    explicitReceiptMonths,
    aggregateTargetMonths,
    aggregateDecisionMonths,
    requestedAggregateMonths,
    missingAggregateMonths,
    aggregateMonthsSource
  };

  return {
    aggregateDecisionMonths,
    isAggregateInvoice: !!(aggregateDecisionMonths && aggregateDecisionMonths.length > 1),
    decisionSources,
    trace
  };
}

function logAggregateDecisionTrace_(label, payload) {
  const logger = typeof billingLogger_ !== 'undefined' && billingLogger_ && typeof billingLogger_.log === 'function'
    ? billingLogger_
    : null;
  if (!logger) return;

  try {
    logger.log('[billing][aggregateDecision][' + label + '] ' + JSON.stringify(payload));
  } catch (err) {
    try {
      logger.log('[billing][aggregateDecision][' + label + ']', payload);
    } catch (e) {
      // ignore logging failures
    }
  }
}

function resolvePreviousBillingMonthKey_(billingMonth) {
  const key = normalizeInvoiceMonthKey_(billingMonth);
  if (!key) return '';

  const year = Number(key.slice(0, 4));
  const month = Number(key.slice(4, 6));
  if (!Number.isFinite(year) || !Number.isFinite(month)) return '';

  const previousMonth = month === 1 ? 12 : month - 1;
  const previousYear = month === 1 ? year - 1 : year;
  return String(previousYear).padStart(4, '0') + String(previousMonth).padStart(2, '0');
}

function isPreviousReceiptSettled_(item) {
  const status = item && (item.previousReceiptStatus || item.receiptStatus);
  return String(status || '').toUpperCase() === 'SETTLED';
}

function buildInvoicePreviousReceipt_(item, display, aggregateMonths) {
  const receiptDisplay = display || resolveInvoiceReceiptDisplay_(item);
  const addressee = item && item.nameKanji ? String(item.nameKanji).trim() : '';
  const receiptMonths = receiptDisplay && receiptDisplay.receiptMonths ? receiptDisplay.receiptMonths : [];
  const aggregateTargetMonths = Array.isArray(item && item.aggregateTargetMonths) ? item.aggregateTargetMonths : [];
  const aggregateMonthsSource = Array.isArray(aggregateMonths) && aggregateMonths.length
    ? aggregateMonths
    : (receiptDisplay && Array.isArray(receiptDisplay.aggregateDecisionMonths) && receiptDisplay.aggregateDecisionMonths.length
      ? receiptDisplay.aggregateDecisionMonths
      : (aggregateTargetMonths.length ? aggregateTargetMonths : receiptMonths));
  const aggregateMonthsForCalc = normalizeAggregateMonthsForInvoice_(
    aggregateMonthsSource,
    item && item.billingMonth
  );
  const monthsForReceipt = aggregateMonthsForCalc.length ? aggregateMonthsForCalc : receiptMonths;
  const date = formatReceiptSettlementDate_(monthsForReceipt, formatInvoiceDateLabel_());
  const receiptMonthsCount = Array.isArray(monthsForReceipt) ? monthsForReceipt.length : 0;
  const isAggregateInvoice = receiptMonthsCount > 1;
  const aggregateBreakdown = isAggregateInvoice ? resolveReceiptMonthBreakdown_(item, aggregateMonthsForCalc) : [];
  const aggregateAmount = Array.isArray(aggregateBreakdown)
    ? aggregateBreakdown.reduce((sum, row) => sum + (Number(row && row.amount) || 0), 0)
    : 0;
  const note = receiptDisplay && receiptDisplay.receiptRemark ? receiptDisplay.receiptRemark : '';

  const resolvedAmount = isAggregateInvoice
    ? aggregateAmount
    : (Array.isArray(aggregateBreakdown) && aggregateBreakdown.length
      ? Number(aggregateBreakdown[0] && aggregateBreakdown[0].amount) || 0
      : 0);

  return {
    visible: !!(receiptDisplay && receiptDisplay.visible),
    addressee,
    date,
    amount: Number.isFinite(resolvedAmount) ? resolvedAmount : 0,
    note,
    receiptMonths: receiptDisplay && receiptDisplay.receiptMonths ? receiptDisplay.receiptMonths : [],
    aggregateStatus: receiptDisplay && receiptDisplay.aggregateStatus ? receiptDisplay.aggregateStatus : '',
    aggregateConfirmed: !!(receiptDisplay && receiptDisplay.aggregateConfirmed),
    settled: isPreviousReceiptSettled_(item)
  };
}

function formatReceiptSettlementDate_(receiptMonths, fallbackDate) {
  const months = Array.isArray(receiptMonths) ? receiptMonths : [];
  const lastMonth = months.length ? normalizeInvoiceMonthKey_(months[months.length - 1]) : '';

  if (lastMonth) {
    const year = Number(lastMonth.slice(0, 4));
    const month = Number(lastMonth.slice(4, 6));
    if (Number.isFinite(year) && Number.isFinite(month)) {
      const settlementDate = new Date(year, month - 1, 20);
      try {
        const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
        return Utilities.formatDate(settlementDate, tz, 'yyyy-MM-20');
      } catch (e) {
        // ignore and fall through
      }
      return settlementDate.toISOString().slice(0, 10);
    }
  }

  return fallbackDate || '';
}

function resolveReceiptMonthBreakdown_(item, receiptMonths) {
  const precomputed = item && Array.isArray(item.receiptMonthBreakdown) ? item.receiptMonthBreakdown : [];
  if (item && item.hasOwnProperty('receiptMonthBreakdown')) return precomputed;

  try {
    return buildReceiptMonthBreakdownForEntry_(
      item && item.patientId,
      receiptMonths || (item && item.receiptMonths) || [],
      item,
      {}
    );
  } catch (e) {
    return [];
  }
}

function formatInvoiceDateLabel_() {
  try {
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    return Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  } catch (e) {
    return '';
  }
}

function sanitizeFileName_(text) {
  const raw = String(text || '').trim();
  return raw ? raw.replace(/[\\/\r\n]/g, '_') : '請求書';
}

function normalizeInvoicePatientIdsForOutput_(patientIds) {
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

function normalizeInvoicePdfContext_(context) {
  const source = context && typeof context === 'object' ? context : {};
  const amount = source.amount && typeof source.amount === 'object' ? source.amount : {};
  const months = Array.isArray(source.months) ? source.months : [];
  return {
    patientId: source.patientId ? String(source.patientId) : '',
    billingMonth: source.billingMonth ? String(source.billingMonth) : '',
    months,
    amount: Object.assign({
      rows: Array.isArray(amount.rows) ? amount.rows : [],
      aggregateMonthTotals: Array.isArray(amount.aggregateMonthTotals) ? amount.aggregateMonthTotals : [],
      aggregateRemark: amount.aggregateRemark || '',
      grandTotal: Number.isFinite(amount.grandTotal) ? amount.grandTotal : 0,
      insuranceType: amount.insuranceType || '',
      burdenRate: amount.burdenRate || '',
      chargeMonthLabel: amount.chargeMonthLabel || '',
      receiptMonths: Array.isArray(amount.receiptMonths) ? amount.receiptMonths : [],
      receiptRemark: amount.receiptRemark || '',
      showReceipt: !!amount.showReceipt,
      previousReceipt: amount.previousReceipt || null,
      forceHideReceipt: !!amount.forceHideReceipt,
      watermark: amount.watermark || null,
      finalized: !!amount.finalized,
      aggregateStatus: amount.aggregateStatus || '',
      aggregateConfirmed: !!amount.aggregateConfirmed
    }, amount),
    name: source.name || source.nameKanji || '',
    isAggregateInvoice: !!source.isAggregateInvoice,
    responsibleName: source.responsibleName || ''
  };
}

function buildInvoiceTemplateContext_(normalizedContext) {
  return {
    patientId: normalizedContext.patientId,
    billingMonth: normalizedContext.billingMonth,
    months: normalizedContext.months,
    amount: normalizedContext.amount,
    name: normalizedContext.name,
    isAggregateInvoice: normalizedContext.isAggregateInvoice
  };
}

function formatInvoiceFileName_(item, options) {
  const baseName = sanitizeFileName_(item && (item.name || item.nameKanji || item.patientId || INVOICE_FILE_PREFIX));
  const dateLabel = formatInvoiceDateLabel_();
  const suffix = options && options.fileNameSuffix ? String(options.fileNameSuffix).trim() : '';
  const base = baseName + '_' + (dateLabel || 'YYYYMMDD') + '_請求書';
  return base + suffix + '.pdf';
}

function formatAggregateInvoiceFileName_(item, billingMonthLabel) {
  const addressee = sanitizeFileName_(item && (item.name || item.nameKanji || item.patientId || INVOICE_FILE_PREFIX));
  const monthLabel = sanitizeFileName_(billingMonthLabel || formatBillingMonthForFile_(item && item.billingMonth));
  return `${INVOICE_FILE_PREFIX}_${addressee}_${monthLabel}（合算）.pdf`;
}

function normalizeAggregateMonthsForInvoice_(months, billingMonth) {
  return normalizePastInvoiceMonths_(months, billingMonth);
}

function formatAggregateInvoiceRemark_(months) {
  if (!Array.isArray(months) || !months.length) return '';
  const labels = months
    .map(ym => normalizeInvoiceMonthKey_(ym))
    .filter(Boolean)
    .map(ym => String(ym).slice(4, 6) + '月');
  if (!labels.length) return '';
  return labels.join('・') + '分 施術料金として';
}

function buildInvoiceWatermark_(item) {
  const finalized = !!(item && item.billingFinalized);
  return finalized ? { text: '確定済み' } : null;
}

function resolveAggregatePreparedBillingEntry_(monthKey, patientId, fallbackItem, monthCache) {
  const normalizedPatientId = patientId ? String(patientId).trim() : '';
  const cache = monthCache || {};
  let entry = null;

  if (normalizedPatientId && typeof getPreparedBillingEntryForMonthCached_ === 'function') {
    try {
      entry = getPreparedBillingEntryForMonthCached_(monthKey, normalizedPatientId, cache);
    } catch (e) {
      entry = null;
    }
  }

  if (!entry && typeof loadPreparedBillingWithSheetFallback_ === 'function') {
    try {
      const prepared = loadPreparedBillingWithSheetFallback_(monthKey, { allowInvalid: true, restoreCache: false });
      const payload = prepared && prepared.prepared !== undefined ? prepared.prepared : prepared;
      const normalizedPrepared = typeof normalizePreparedBilling_ === 'function'
        ? normalizePreparedBilling_(payload)
        : payload;
      const billingEntries = normalizedPrepared && normalizedPrepared.billingJson;
      if (Array.isArray(billingEntries)) {
        entry = billingEntries.find(row => {
          const pid = typeof billingNormalizePatientId_ === 'function'
            ? billingNormalizePatientId_(row && row.patientId)
            : (row && row.patientId ? String(row.patientId).trim() : '');
          return pid && pid === normalizedPatientId;
        });
      }
    } catch (err) {
      entry = null;
    }
  }

  if (!entry && fallbackItem && normalizeInvoiceMonthKey_(fallbackItem.billingMonth) === monthKey) {
    entry = fallbackItem;
  }

  return entry ? Object.assign({ billingMonth: monthKey }, entry) : null;
}

function filterAggregateMonthsByPrepared_(item, aggregateMonths, options) {
  const billingMonth = item && item.billingMonth;
  const normalizedPatientId = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(item && item.patientId)
    : (item && item.patientId ? String(item.patientId).trim() : '');
  const monthCache = options && options.monthCache ? options.monthCache : {};
  const months = normalizeAggregateMonthsForInvoice_(aggregateMonths, billingMonth);
  const availableMonths = [];
  const missingMonths = [];

  months.forEach(monthKey => {
    const entry = resolveAggregatePreparedBillingEntry_(monthKey, normalizedPatientId, item, monthCache);
    if (entry) {
      availableMonths.push(monthKey);
    } else {
      missingMonths.push(monthKey);
    }
  });

  return { months: availableMonths, missingMonths, monthCache };
}

function buildAggregateInvoiceBreakdowns_(item, aggregateMonths, options) {
  const billingMonth = item && item.billingMonth;
  const months = normalizeAggregateMonthsForInvoice_(aggregateMonths, billingMonth);
  if (!months.length) {
    return {
      months: [],
      breakdowns: [],
      totals: null,
      unitPriceSource: null,
      hasManualTransport: false
    };
  }

  const normalizedPatientId = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(item && item.patientId)
    : (item && item.patientId ? String(item.patientId).trim() : '');
  const monthCache = options && options.monthCache ? options.monthCache : {};

  const breakdowns = months.map(monthKey => {
    const entry = resolveAggregatePreparedBillingEntry_(monthKey, normalizedPatientId, item, monthCache);
    if (!entry) {
      return { month: monthKey, entry: null, breakdown: null, carryOverAmount: 0 };
    }
    const breakdown = calculateInvoiceChargeBreakdown_(Object.assign({}, entry, { billingMonth: monthKey }));
    const carryOverAmount = normalizeBillingCarryOver_(entry);
    return { month: monthKey, entry, breakdown, carryOverAmount };
  });

  const totals = breakdowns.reduce((sum, row) => {
    const breakdown = row.breakdown || {};
    sum.visits += breakdown.visits || 0;
    sum.treatmentAmount += Number(breakdown.treatmentAmount) || 0;
    sum.transportAmount += Number(breakdown.transportAmount) || 0;
    sum.selfPayTotal += Number(breakdown.selfPayTotal) || 0;
    sum.grandTotal += Number(breakdown.grandTotal) || 0;
    sum.carryOverAmount += Number(row.carryOverAmount) || 0;
    return sum;
  }, {
    visits: 0,
    treatmentAmount: 0,
    transportAmount: 0,
    selfPayTotal: 0,
    grandTotal: 0,
    carryOverAmount: 0
  });

  const unitPriceSource = breakdowns.find(row => Number.isFinite(row.breakdown && row.breakdown.treatmentUnitPrice));
  const hasManualTransport = breakdowns.some(row => row.breakdown && row.breakdown.transportDetail === '手動入力');

  return {
    months,
    breakdowns,
    totals,
    unitPriceSource,
    hasManualTransport
  };
}

function buildAggregateInvoiceTemplateData_(item, aggregateMonths) {
  const billingMonth = item && item.billingMonth;
  const normalizedBillingMonth = normalizeInvoiceMonthKey_(billingMonth);
  const monthLabel = normalizeBillingMonthLabel_(billingMonth);
  const aggregateLabel = monthLabel ? `${monthLabel}（合算）` : '合算請求';
  const baseAmount = normalizeBillingAmount_(item);
  const watermark = buildInvoiceWatermark_(item);
  const initialReceipt = resolveInvoiceReceiptDisplay_(item);
  const receiptDecisionMonths = initialReceipt && Array.isArray(initialReceipt.explicitReceiptMonths)
    ? initialReceipt.explicitReceiptMonths
    : [];
  const months = normalizeAggregateMonthsForInvoice_(
    Array.isArray(aggregateMonths) && aggregateMonths.length
      ? aggregateMonths
    : (receiptDecisionMonths.length
      ? receiptDecisionMonths
      : (Array.isArray(item && item.aggregateTargetMonths) ? item.aggregateTargetMonths : [])),
    billingMonth
  );
  const monthCache = {};
  const preparedAggregate = filterAggregateMonthsByPrepared_(item, months, { monthCache });
  const filteredMonths = preparedAggregate.months;
  const missingAggregateMonths = preparedAggregate.missingMonths;
  const receipt = resolveInvoiceReceiptDisplay_(item, { aggregateMonths: filteredMonths });
  const receiptMonths = receipt && receipt.receiptMonths ? receipt.receiptMonths : [];
  const aggregateRemark = formatAggregateInvoiceRemark_(filteredMonths);
  const aggregateData = buildAggregateInvoiceBreakdowns_(item, filteredMonths, { monthCache });
  const aggregateMonthTotals = aggregateData.breakdowns.map(row => ({
    month: row.month,
    monthLabel: normalizeBillingMonthLabel_(row.month),
    total: row.breakdown ? row.breakdown.grandTotal : 0
  }));
  const aggregateTotal = aggregateMonthTotals.length
    ? aggregateData.totals.grandTotal
    : baseAmount;
  const amount = aggregateTotal;
  const normalizedPatientId = typeof billingNormalizePatientId_ === 'function'
    ? billingNormalizePatientId_(item && item.patientId)
    : String(item && item.patientId || '').trim();
  const aggregateStatus = receipt ? receipt.aggregateStatus : normalizeAggregateStatus_(item && item.aggregateStatus);
  const aggregateConfirmed = receipt ? receipt.aggregateConfirmed : aggregateStatus === 'confirmed';
  const basePreviousReceipt = buildInvoicePreviousReceipt_(item, receipt, months);
  const previousReceipt = item && item.previousReceipt
    ? Object.assign({}, basePreviousReceipt, item.previousReceipt)
    : basePreviousReceipt;
  if (previousReceipt) {
    previousReceipt.settled = isPreviousReceiptSettled_(item);
  }

  if (previousReceipt) {
    previousReceipt.visible = !!(receipt && receipt.visible);
  }

  const aggregateDecision = resolveAggregateInvoiceDecision_(item, receipt, billingMonth, { monthCache });
  const aggregateDecisionTrace = Object.assign(
    { decisionSources: aggregateDecision && aggregateDecision.decisionSources ? aggregateDecision.decisionSources : [] },
    aggregateDecision && aggregateDecision.trace ? aggregateDecision.trace : {},
    { isAggregateInvoice: true, aggregateMonthsForOutput: filteredMonths, missingAggregateMonths }
  );
  logAggregateDecisionTrace_('aggregate_template', Object.assign(
    {
      patientId: normalizedPatientId,
      nameKanji: item && item.nameKanji ? String(item.nameKanji) : ''
    },
    aggregateDecisionTrace
  ));

  return Object.assign({}, item, {
    monthLabel: aggregateLabel,
    chargeMonthLabel: monthLabel,
    isAggregateInvoice: true,
    invoiceMode: 'aggregate',
    watermark,
    aggregateStatus,
    aggregateConfirmed,
    receiptMonths: filteredMonths,
    receiptRemark: aggregateRemark,
    aggregateRemark,
    aggregateMonthTotals,
    showReceipt: !!(receipt && receipt.visible),
    rows: [
      { label: '合算請求額', detail: aggregateRemark, amount }
    ],
    grandTotal: aggregateTotal,
    previousReceipt,
    finalized: !!(aggregateConfirmed || (previousReceipt && previousReceipt.settled)),
    aggregateDecisionTrace
  });
}

function buildInvoiceTemplateData_(item) {
  const billingMonth = item && item.billingMonth;
  const monthLabel = normalizeBillingMonthLabel_(billingMonth);
  const watermark = buildInvoiceWatermark_(item);

  const initialReceipt = resolveInvoiceReceiptDisplay_(item);
  const aggregateStatus = initialReceipt ? initialReceipt.aggregateStatus : normalizeAggregateStatus_(item && item.aggregateStatus);
  const aggregateConfirmed = initialReceipt ? initialReceipt.aggregateConfirmed : aggregateStatus === 'confirmed';
  const monthCache = {};
  const aggregateDecision = resolveAggregateInvoiceDecision_(item, initialReceipt, billingMonth, { monthCache });
  const aggregateDecisionMonths = aggregateDecision && aggregateDecision.aggregateDecisionMonths
    ? aggregateDecision.aggregateDecisionMonths
    : [];
  const receipt = resolveInvoiceReceiptDisplay_(item, { aggregateMonths: aggregateDecisionMonths });
  const receiptMonths = receipt && receipt.receiptMonths ? receipt.receiptMonths : [];
  const basePreviousReceipt = buildInvoicePreviousReceipt_(item, receipt, aggregateDecisionMonths);
  const previousReceipt = item && item.previousReceipt
    ? Object.assign({}, basePreviousReceipt, item.previousReceipt)
    : basePreviousReceipt;
  if (previousReceipt) {
    previousReceipt.settled = isPreviousReceiptSettled_(item);
  }

  if (previousReceipt) {
    previousReceipt.visible = !!(receipt && receipt.visible);
  }

  const aggregateDecisionTrace = Object.assign(
    { decisionSources: aggregateDecision && aggregateDecision.decisionSources ? aggregateDecision.decisionSources : [] },
    aggregateDecision && aggregateDecision.trace ? aggregateDecision.trace : {},
    { isAggregateInvoice: aggregateDecision && aggregateDecision.isAggregateInvoice }
  );
  logAggregateDecisionTrace_('invoice_template', Object.assign(
    {
      patientId: item && item.patientId ? String(item.patientId) : '',
      nameKanji: item && item.nameKanji ? String(item.nameKanji) : ''
    },
    aggregateDecisionTrace
  ));
  const isAggregateInvoice = !!(aggregateDecision && aggregateDecision.isAggregateInvoice);
  const chargeMonthLabel = monthLabel;
  let breakdown = calculateInvoiceChargeBreakdown_(Object.assign({}, item, { billingMonth }));
  let visits = breakdown.visits || 0;
  let unitPrice = breakdown.treatmentUnitPrice || 0;
  let carryOverAmount = normalizeBillingCarryOver_(item);
  let transportDetail = breakdown.transportDetail || (formatBillingCurrency_(TRANSPORT_PRICE) + '円 × ' + visits + '回');
  let aggregateMonthTotals = [];
  let grandTotal = breakdown.grandTotal;

  if (isAggregateInvoice && aggregateDecisionMonths.length > 1) {
    const aggregateData = buildAggregateInvoiceBreakdowns_(item, aggregateDecisionMonths, { monthCache });
    const totals = aggregateData.totals || {};
    visits = totals.visits || 0;
    carryOverAmount = totals.carryOverAmount || 0;
    unitPrice = aggregateData.unitPriceSource ? aggregateData.unitPriceSource.breakdown.treatmentUnitPrice || 0 : unitPrice;
    breakdown = Object.assign({}, breakdown, {
      visits,
      treatmentAmount: totals.treatmentAmount || 0,
      transportAmount: totals.transportAmount || 0,
      selfPayTotal: totals.selfPayTotal || 0,
      grandTotal: totals.grandTotal || 0
    });
    transportDetail = aggregateData.hasManualTransport
      ? '手動入力'
      : formatBillingCurrency_(TRANSPORT_PRICE) + '円 × ' + visits + '回';
    aggregateMonthTotals = aggregateData.breakdowns.map(row => ({
      month: row.month,
      monthLabel: normalizeBillingMonthLabel_(row.month),
      total: row.breakdown ? row.breakdown.grandTotal : 0
    }));
    grandTotal = breakdown.grandTotal;
  }

  const rows = [
    { label: '前月繰越', detail: '', amount: carryOverAmount },
    { label: '施術料', detail: formatBillingCurrency_(unitPrice) + '円 × ' + visits + '回', amount: breakdown.treatmentAmount },
    { label: '交通費', detail: transportDetail, amount: breakdown.transportAmount }
  ];

  return Object.assign({}, item, {
    monthLabel,
    chargeMonthLabel,
    isAggregateInvoice,
    invoiceMode: isAggregateInvoice ? 'aggregate' : 'standard',
    watermark,
    aggregateStatus,
    aggregateConfirmed,
    receiptMonths,
    receiptRemark: (receipt && receipt.receiptRemark) || '',
    showReceipt: !!(receipt && receipt.visible),
    rows,
    grandTotal,
    previousReceipt,
    finalized: !!(aggregateConfirmed || (previousReceipt && previousReceipt.settled)),
    aggregateMonthTotals,
    aggregateDecisionTrace
  });
}

function createInvoicePdfBlob_(context, options) {
  const template = HtmlService.createTemplateFromFile('invoice_template');
  const normalizedContext = normalizeInvoicePdfContext_(context || {});
  template.data = buildInvoiceTemplateContext_(normalizedContext);
  const html = template.evaluate().setWidth(1240).setHeight(1754);
  const fileName = formatInvoiceFileName_(normalizedContext, options);
  return html.getBlob().getAs(MimeType.PDF).setName(fileName);
}

function createAggregateInvoicePdfBlob_(context, options) {
  const normalizedContext = normalizeInvoicePdfContext_(context || {});
  const template = HtmlService.createTemplateFromFile('invoice_template');
  template.data = buildInvoiceTemplateContext_(normalizedContext);
  const html = template.evaluate().setWidth(1240).setHeight(1754);
  const billingLabel = normalizedContext.billingMonth ? normalizeBillingMonthLabel_(normalizedContext.billingMonth) : '';
  const aggregateLabel = billingLabel ? `${billingLabel}（合算）` : '合算請求';
  const fileName = formatAggregateInvoiceFileName_(
    normalizedContext,
    aggregateLabel.replace(/請求書?\s*/g, '')
  );
  return html.getBlob().getAs(MimeType.PDF).setName(fileName);
}

function ensureInvoiceRootFolder_() {
  if (!INVOICE_PARENT_FOLDER_ID) {
    throw new Error('請求書の保存先フォルダIDが設定されていません');
  }
  return DriveApp.getFolderById(INVOICE_PARENT_FOLDER_ID);
}

function ensureSubFolder_(parentFolder, name) {
  const folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(name);
}

function formatResponsibleFolderName_(billingMonth, responsibleName) {
  const ym = formatBillingMonthCompact_(billingMonth);
  const ymLabel = ym || '請求月未設定';
  const safeName = sanitizeFileName_(responsibleName || '担当者未設定');
  return ymLabel + '請求書_' + safeName;
}

function calculateInvoiceChargeBreakdown_(params) {
  const visits = normalizeInvoiceVisitCount_(params && params.visitCount);
  const insuranceType = params && params.insuranceType ? String(params.insuranceType).trim() : '';
  const burdenRateInt = normalizeInvoiceBurdenRateInt_(params && params.burdenRate);
  const normalizedMedicalAssistance = normalizeInvoiceMedicalAssistanceFlag_(params && params.medicalAssistance);
  const carryOverAmount = normalizeBillingCarryOver_(params);
  const manualUnitPrice = params && params.hasOwnProperty('manualUnitPrice')
    ? params.manualUnitPrice
    : params && params.unitPrice;
  const hasManualTransportInput = params && Object.prototype.hasOwnProperty.call(params, 'manualTransportAmount');
  const manualTransportInput = hasManualTransportInput ? params.manualTransportAmount : null;
  const manualTransportAmount = (manualTransportInput === ''
    || manualTransportInput === null
    || manualTransportInput === undefined
    || !hasManualTransportInput)
    ? null
    : normalizeInvoiceMoney_(manualTransportInput);
  const patientUnitPrice = params && params.unitPrice;
  const treatmentUnitPrice = resolveInvoiceUnitPriceForOutput_(
    insuranceType,
    burdenRateInt,
    manualUnitPrice,
    normalizedMedicalAssistance,
    patientUnitPrice || INVOICE_TREATMENT_UNIT_PRICE_BY_BURDEN[burdenRateInt]
  );

  const hasChargeableUnitPrice = Number.isFinite(treatmentUnitPrice) && treatmentUnitPrice !== 0;
  const treatmentAmountFull = visits > 0 && hasChargeableUnitPrice ? treatmentUnitPrice * visits : 0;
  const isSelfPaid = insuranceType === '自費' || burdenRateInt === '自費';
  const defaultBurdenUnitPrice = INVOICE_TREATMENT_UNIT_PRICE_BY_BURDEN[burdenRateInt];
  const usesBurdenAdjustedUnitPrice = Number.isFinite(defaultBurdenUnitPrice)
    && treatmentUnitPrice === defaultBurdenUnitPrice;
  const burdenMultiplier = isSelfPaid || usesBurdenAdjustedUnitPrice
    ? 1
    : (typeof normalizeBurdenMultiplier_ === 'function'
      ? normalizeBurdenMultiplier_(burdenRateInt, insuranceType)
      : (insuranceType === '自費' ? 1 : (burdenRateInt > 0 ? burdenRateInt / 10 : 0)));
  const treatmentAmount = isSelfPaid
    ? treatmentAmountFull
    : treatmentAmountFull * burdenMultiplier;
  const transportAmount = (manualTransportInput !== '' && manualTransportInput !== null && manualTransportInput !== undefined
    && Number.isFinite(manualTransportAmount))
    ? manualTransportAmount
    : visits > 0 && hasChargeableUnitPrice ? TRANSPORT_PRICE * visits : 0;
  const transportDetail = (manualTransportInput !== '' && manualTransportInput !== null && manualTransportInput !== undefined
    && hasManualTransportInput)
    ? '手動入力'
    : formatBillingCurrency_(TRANSPORT_PRICE) + '円 × ' + visits + '回';
  const selfPayItems = Array.isArray(params && params.selfPayItems)
    ? params.selfPayItems
    : (params && params.manualSelfPayAmount ? [{ type: '自費', amount: params.manualSelfPayAmount }] : []);
  const selfPayTotal = selfPayItems.reduce((sum, entry) => sum + (normalizeInvoiceMoney_(entry.amount) || 0), 0);
  const grandTotal = carryOverAmount + treatmentAmount + transportAmount + selfPayTotal;

  return { treatmentUnitPrice, treatmentAmount, transportAmount, transportDetail, grandTotal, visits, selfPayItems, selfPayTotal };
}

function buildBillingInvoiceHtml_(item, billingMonth) {
  const targetMonth = billingMonth || (item && item.billingMonth) || '';
  let breakdown = calculateInvoiceChargeBreakdown_(Object.assign({}, item, { billingMonth: targetMonth }));
  const monthLabel = normalizeBillingMonthLabel_(targetMonth);
  const aggregateMonths = Array.isArray(item && item.aggregateTargetMonths)
    ? normalizeAggregateMonthsForInvoice_(item.aggregateTargetMonths, targetMonth)
    : [];
  const preparedAggregate = filterAggregateMonthsByPrepared_(item, aggregateMonths, {});
  const filteredAggregateMonths = preparedAggregate.months;
  const transportUnitPrice = TRANSPORT_PRICE;
  let carryOverAmount = normalizeBillingCarryOver_(item);
  let visits = breakdown.visits || 0;
  let treatmentUnitPrice = breakdown.treatmentUnitPrice || 0;
  let treatmentAmount = breakdown.treatmentAmount || 0;
  let transportAmount = breakdown.transportAmount || 0;
  let selfPayTotal = breakdown.selfPayTotal || 0;
  let transportDetail = breakdown.transportDetail || (formatBillingCurrency_(transportUnitPrice) + '円 × ' + visits + '回');
  let grandTotal = breakdown.grandTotal;

  if (filteredAggregateMonths.length) {
    const aggregateData = buildAggregateInvoiceBreakdowns_(item, filteredAggregateMonths, {});
    if (aggregateData && aggregateData.breakdowns.length) {
      const totals = aggregateData.totals || {};
      visits = totals.visits || 0;
      carryOverAmount = totals.carryOverAmount || 0;
      treatmentAmount = totals.treatmentAmount || 0;
      transportAmount = totals.transportAmount || 0;
      selfPayTotal = totals.selfPayTotal || 0;
      const unitPriceSource = aggregateData.unitPriceSource;
      if (unitPriceSource) {
        treatmentUnitPrice = unitPriceSource.breakdown.treatmentUnitPrice || 0;
      }
      transportDetail = aggregateData.hasManualTransport
        ? '手動入力'
        : formatBillingCurrency_(transportUnitPrice) + '円 × ' + visits + '回';
      grandTotal = totals.grandTotal || 0;
    }
  }

  const totalLabel = formatBillingCurrency_(grandTotal) + '円';

  const name = escapeHtml_((item && item.nameKanji) || '');

  return [
    '<div class="billing-invoice">',
    '<h1>べるつりー訪問鍼灸マッサージ</h1>',
    `<h2>${escapeHtml_(monthLabel)} ご請求書</h2>`,
    name ? `<p class="patient-name">${name} 様</p>` : '',
    '<div class="charge-breakdown">',
    `<p>前月繰越: ${formatBillingCurrency_(carryOverAmount)}円</p>`,
    `<p>施術料（${formatBillingCurrency_(treatmentUnitPrice)}円 × ${visits}回）: ${formatBillingCurrency_(treatmentAmount)}円</p>`,
    `<p>交通費（${transportDetail}）: ${formatBillingCurrency_(transportAmount)}円</p>`,
    `<p class="grand-total">合計: ${totalLabel}</p>`,
    '</div>',
    '</div>'
  ].filter(Boolean).join('');
}

function ensureInvoiceFolderForResponsible_(item) {
  const root = ensureInvoiceRootFolder_();
  const folderName = formatResponsibleFolderName_(item && item.billingMonth, item && item.responsibleName);
  return ensureSubFolder_(root, folderName);
}

function removeExistingInvoiceFiles_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    const file = files.next();
    try {
      file.setTrashed(true);
    } catch (e) {
      // ignore cleanup errors
    }
  }
}

function saveInvoicePdf(item, pdfBlob, options) {
  const folder = ensureInvoiceFolderForResponsible_(item);
  const shouldOverwrite = !(options && options.overwriteExisting === false);
  const fileName = pdfBlob.getName();
  if (shouldOverwrite) {
    removeExistingInvoiceFiles_(folder, fileName);
  }
  const file = folder.createFile(pdfBlob);
  return { fileId: file.getId(), url: file.getUrl(), name: file.getName() };
}

function generateInvoicePdf(context, options) {
  const normalizedContext = normalizeInvoicePdfContext_(context || {});
  const blob = normalizedContext.isAggregateInvoice
    ? createAggregateInvoicePdfBlob_(normalizedContext, options)
    : createInvoicePdfBlob_(normalizedContext, options);
  return saveInvoicePdf(normalizedContext, blob, options);
}

function generateAggregateInvoicePdf(context, options) {
  const normalizedContext = normalizeInvoicePdfContext_(context || {});
  const blob = createAggregateInvoicePdfBlob_(normalizedContext, options);
  return saveInvoicePdf(normalizedContext, blob, options);
}

function generateInvoicePdfs(invoiceContexts, options) {
  const list = Array.isArray(invoiceContexts) ? invoiceContexts : [];
  const billingMonth = (options && options.billingMonth) || (list.length && list[0].billingMonth) || '';
  const patientIds = normalizeInvoicePatientIdsForOutput_(options && options.patientIds);
  const targets = patientIds.length
    ? list.filter(item => patientIds.indexOf(String(item && item.patientId ? item.patientId : '').trim()) >= 0)
    : list;
  const isPartialGeneration = patientIds.length > 0;
  const invoiceFileOptions = {
    overwriteExisting: !isPartialGeneration,
    fileNameSuffix: isPartialGeneration ? (options && options.reissueSuffix ? options.reissueSuffix : '_再発行') : ''
  };
  const files = targets.map(item => {
    const meta = generateInvoicePdf(item, invoiceFileOptions);
    return Object.assign({}, meta, { patientId: item && item.patientId, nameKanji: item && (item.name || item.nameKanji) });
  });
  const matchedIds = new Set(files.map(f => String(f.patientId || '').trim()).filter(Boolean));
  const missingPatientIds = patientIds.filter(id => !matchedIds.has(id));
  return { billingMonth, files, missingPatientIds, requestedPatientIds: patientIds };
}

/***** Bank transfer export helpers (legacy; not used in new billing specification) *****/

const LEGACY_BANK_TRANSFER_SHEET_NAME = '銀行データ出力（旧仕様）';
const BANK_TRANSFER_HEADERS = ['請求月', '番号', '氏名（漢字）', '銀行コード', '支店コード', '規定コード', '口座番号', '氏名（カナ）', '新規フラグ'];

function buildBankTransferRowsForBilling_(billingJson, bankInfoByName, patientMap, billingMonth, bankStatuses) {
  const rows = [];
  let skipped = 0;
  let passed = 0;
  let total = 0;
  const skipReasons = {
    invalidBankCode: 0,
    invalidBranchCode: 0,
    invalidAccountNumber: 0
  };
  const billingMonthKey = billingMonth || (billingJson && billingJson.length ? billingJson[0].billingMonth : '');

  const billedByPatientId = (billingJson || []).reduce((map, item) => {
    const pid = item && item.patientId ? String(item.patientId).trim() : '';
    if (pid && !map[pid]) {
      map[pid] = item;
    }
    return map;
  }, {});

  const patientsByNameKey = Object.values(patientMap || {}).reduce((map, patient) => {
    if (!patient) return map;
    const nameKey = buildBillingNameKey_(patient);
    if (nameKey && !map[nameKey]) {
      map[nameKey] = patient;
    }
    return map;
  }, {});

  const patientByIdMap = patientMap || {};
  const bankEntryMap = Object.entries(bankInfoByName || {})
    .filter(([, entry]) => !!entry)
    .reduce((map, [nameKey, entry]) => {
      const normalizedKey = buildBillingNameKey_(Object.assign({ _nameKey: nameKey }, entry));
      if (normalizedKey && !map[normalizedKey]) {
        map[normalizedKey] = Object.assign({ _nameKey: normalizedKey }, entry);
      }
      return map;
    }, {});

  const billedByNameKey = (billingJson || []).reduce((map, item) => {
    const nameKey = buildBillingNameKey_(item);
    if (nameKey && !map[nameKey]) {
      map[nameKey] = item;
    }
    return map;
  }, {});

  (billingJson || []).forEach(item => {
    if (!item) return;
    const pid = item.patientId ? String(item.patientId).trim() : '';
    const nameKey = buildBillingNameKey_(item);
    if (!nameKey || bankEntryMap[nameKey]) return;
    const patient = pid && patientByIdMap[pid] ? patientByIdMap[pid] : {};
    bankEntryMap[nameKey] = Object.assign({ _nameKey: nameKey }, patient, item);
  });

  const bankEntries = Object.values(bankEntryMap);
  const seenNameKeys = new Set();

  bankEntries.forEach(bankEntry => {
    const nameKey = buildBillingNameKey_(bankEntry);
    if (!nameKey || seenNameKeys.has(nameKey)) return;
    seenNameKeys.add(nameKey);

    const patientIdFromEntry = bankEntry.patientId ? String(bankEntry.patientId).trim() : '';
    const patient = patientsByNameKey[nameKey]
      || (patientIdFromEntry && patientByIdMap[patientIdFromEntry])
      || {};
    const item = (() => {
      const pidKey = patient && patient.patientId ? String(patient.patientId).trim() : patientIdFromEntry;
      if (pidKey && billedByPatientId[pidKey]) {
        return billedByPatientId[pidKey];
      }
      return billedByNameKey[nameKey] || null;
    })();
    const pid = item && item.patientId
      ? String(item.patientId).trim()
      : (patient && patient.patientId ? String(patient.patientId).trim() : '');
    if (!item) return;

    total += 1;

    const pickWithPriority = (resolver, fallbackValue) => {
      const sources = [bankEntry, patient, item];
      for (let i = 0; i < sources.length; i += 1) {
        const source = sources[i];
        if (!source) continue;
        const value = typeof resolver === 'function' ? resolver(source) : source[resolver];
        if (value !== undefined && value !== null && String(value).trim() !== '') {
          return value;
        }
      }
      return fallbackValue;
    };

    const rawBankCode = pickWithPriority('bankCode', '');
    const bankCodeDigits = String(rawBankCode).replace(/\D/g, '');
    const bankCode = bankCodeDigits ? bankCodeDigits.padStart(4, '0') : '';
    const rawBranchCode = pickWithPriority('branchCode', '');
    const branchCodeDigits = String(rawBranchCode).replace(/\D/g, '');
    const branchCode = branchCodeDigits ? branchCodeDigits.padStart(3, '0') : '';
    const regulationCode = pickWithPriority('regulationCode', 1);
    const mergedNameKanji = pickWithPriority('nameKanji', bankEntry.nameKanji);
    const rawNameKana = pickWithPriority('nameKana', bankEntry.nameKana);
    const nameKana = rawNameKana ? normalizeKana_(rawNameKana) : normalizeKana_(mergedNameKanji);
    const rawAccountNumber = pickWithPriority('accountNumber', '');
    const accountNumberDigits = String(rawAccountNumber).replace(/\D/g, '');
    const accountNumber = accountNumberDigits ? accountNumberDigits.padStart(7, '0') : '';
    const isNew = normalizeZeroOneFlag_(pickWithPriority('isNew', ''));
    const statusEntry = bankStatuses && pid ? bankStatuses[pid] : null;
    const paidStatus = item && item.paidStatus ? item.paidStatus : (statusEntry && statusEntry.paidStatus ? statusEntry.paidStatus : '');

    const bankCodeInvalid = !bankCodeDigits || bankCode.length !== 4;
    const branchCodeInvalid = !branchCodeDigits || branchCode.length !== 3;
    const accountNumberInvalid = !accountNumberDigits || accountNumber.length !== 7;

    if (bankCodeInvalid || branchCodeInvalid || accountNumberInvalid) {
      skipped += 1;
      if (bankCodeInvalid) skipReasons.invalidBankCode += 1;
      if (branchCodeInvalid) skipReasons.invalidBranchCode += 1;
      if (accountNumberInvalid) skipReasons.invalidAccountNumber += 1;
      return;
    }

    passed += 1;
    rows.push({
      billingMonth: billingMonthKey,
      patientId: pid,
      nameKanji: mergedNameKanji,
      bankCode,
      branchCode,
      regulationCode,
      accountNumber,
      nameKana,
      isNew,
      paidStatus
    });
  });

  return {
    billingMonth: billingMonthKey,
    rows,
    skipped,
    total,
    passed,
    skipReasons
  };
}

function ensureBankTransferSheet_() {
  const workbook = resolveBillingSpreadsheetForOutput_();
  let sheet = workbook.getSheetByName(LEGACY_BANK_TRANSFER_SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(LEGACY_BANK_TRANSFER_SHEET_NAME);
    sheet.getRange(1, 1, 1, BANK_TRANSFER_HEADERS.length).setValues([BANK_TRANSFER_HEADERS]);
    return { sheet, headers: BANK_TRANSFER_HEADERS.slice() };
  }

  const lastCol = Math.max(sheet.getLastColumn(), BANK_TRANSFER_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  return { sheet, headers };
}

function resolveBankTransferColumns_(sheet, headers) {
  const workingHeaders = headers ? headers.slice() : [];
  const resolved = {};

  function ensureColumn(label, candidates) {
    const idx = resolveBillingColumn_(workingHeaders, candidates, label, {});
    if (idx) {
      resolved[label] = idx;
      return idx;
    }
    const newIndex = workingHeaders.length + 1;
    sheet.getRange(1, newIndex).setValue(label);
    workingHeaders.push(label);
    resolved[label] = newIndex;
    return newIndex;
  }

  ensureColumn('請求月', ['請求月', 'billingMonth', '請求年月']);
  ensureColumn('番号', BILLING_LABELS.recNo.concat(['番号', '患者番号', '患者ID']));
  ensureColumn('氏名（漢字）', BILLING_LABELS.name.concat(['氏名', '氏名（漢字）']));
  ensureColumn('銀行コード', ['銀行コード', '銀行CD', '銀行番号', 'bankCode']);
  ensureColumn('支店コード', ['支店コード', '支店番号', '支店CD', 'branchCode']);
  ensureColumn('規定コード', ['規定コード', '規定', '規定CD', '規定コード(1固定)', '1固定']);
  ensureColumn('口座番号', ['口座番号', '口座No', '口座NO', 'accountNumber', '口座']);
  ensureColumn('氏名（カナ）', BILLING_LABELS.furigana.concat(['氏名（カナ）']));
  ensureColumn('新規フラグ', ['新規', '新患', 'isNew', '新規フラグ', '新規区分']);
  ensureColumn('領収状態', ['領収状態', '領収', 'paidStatus']);

  return { columns: resolved, headers: workingHeaders };
}

function exportBankTransferRows_(billingMonth, rowObjects, bankStatuses) {
  const ensured = ensureBankTransferSheet_();
  const sheet = ensured.sheet;
  const { columns, headers } = resolveBankTransferColumns_(sheet, ensured.headers);
  const colCount = Math.max(sheet.getLastColumn(), headers.length, Math.max.apply(null, Object.values(columns)));

  const lastRow = sheet.getLastRow();
  const existingValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, colCount).getValues() : [];
  const keyForRow = (row) => {
    const month = columns['請求月'] ? String(row[columns['請求月'] - 1] || '').trim() : '';
    const pid = columns['番号'] ? String(row[columns['番号'] - 1] || '').trim() : '';
    return month && pid ? `${month}::${pid}` : '';
  };

  const workingRowsByKey = existingValues.reduce((map, row) => {
    const key = keyForRow(row);
    if (key && !map.has(key)) map.set(key, row);
    return map;
  }, new Map());

  const mapped = (rowObjects || []).map(obj => {
    const row = new Array(colCount).fill('');
    row[columns['請求月'] - 1] = billingMonth;
    row[columns['番号'] - 1] = obj.patientId || '';
    row[columns['氏名（漢字）'] - 1] = obj.nameKanji || '';
    row[columns['銀行コード'] - 1] = obj.bankCode || '';
    row[columns['支店コード'] - 1] = obj.branchCode || '';
    row[columns['規定コード'] - 1] = obj.regulationCode || '';
    row[columns['口座番号'] - 1] = obj.accountNumber || '';
    row[columns['氏名（カナ）'] - 1] = obj.nameKana || '';
    row[columns['新規フラグ'] - 1] = obj.isNew || '';
    if (columns['領収状態']) {
      const statusEntry = bankStatuses && obj.patientId ? bankStatuses[obj.patientId] : null;
      const existingPaid = row[columns['領収状態'] - 1] || '';
      const paidStatus = (obj.paidStatus != null && obj.paidStatus !== '')
        ? obj.paidStatus
        : (statusEntry && statusEntry.paidStatus ? statusEntry.paidStatus : existingPaid);
      row[columns['領収状態'] - 1] = paidStatus || '';
    }
    return row;
  });

  mapped.forEach(row => {
    const key = keyForRow(row);
    if (key) {
      const existingRow = workingRowsByKey.get(key);
      const mergedRow = existingRow ? existingRow.slice() : new Array(colCount).fill('');
      ['請求月', '番号', '氏名（漢字）', '銀行コード', '支店コード', '規定コード', '口座番号', '氏名（カナ）', '新規フラグ']
        .forEach(label => {
          if (columns[label]) {
            mergedRow[columns[label] - 1] = row[columns[label] - 1];
          }
        });
      const existingPaidStatus = columns['領収状態'] ? mergedRow[columns['領収状態'] - 1] : '';
      if (columns['領収状態']) {
        const statusEntry = bankStatuses && row[columns['番号'] - 1] ? bankStatuses[row[columns['番号'] - 1]] : null;
        const resolvedPaidStatus = (row[columns['領収状態'] - 1] != null && row[columns['領収状態'] - 1] !== '')
          ? row[columns['領収状態'] - 1]
          : (statusEntry && statusEntry.paidStatus ? statusEntry.paidStatus : existingPaidStatus || '');
        mergedRow[columns['領収状態'] - 1] = resolvedPaidStatus;
      }
      workingRowsByKey.set(key, mergedRow);
    }
  });

  const sortedKeys = Array.from(workingRowsByKey.keys()).sort();
  const workingRows = sortedKeys.map(key => workingRowsByKey.get(key));
  const dataRowCount = Math.max(0, lastRow - 1);
  const maxRowCount = Math.max(dataRowCount, workingRows.length);

  if (maxRowCount > 0) {
    sheet.getRange(2, 1, maxRowCount, colCount).clearContent();
  }
  if (workingRows.length) {
    sheet.getRange(2, 1, workingRows.length, colCount).setValues(workingRows);
  }

  return { billingMonth, inserted: mapped.length };
}

  function logPreparedBankPayloadStatus_(prepared) {
    const requiredKeys = ['billingJson', 'visitsByPatient', 'totalsByPatient', 'carryOverByPatient', 'unpaidHistory', 'bankAccountInfoByPatient'];
    const normalized = normalizePreparedBilling_(prepared) || {};
    const preparedWithUnpaidHistory = Object.assign({}, normalized, {
      unpaidHistory: Array.isArray(normalized.unpaidHistory) ? normalized.unpaidHistory : []
    });
    const missing = requiredKeys.filter(key => {
      const value = preparedWithUnpaidHistory[key];
      if (key === 'billingJson') {
        return !Array.isArray(value) || value.length === 0;
      }
      if (key === 'unpaidHistory') {
        return !Array.isArray(value);
      }
      if (Array.isArray(value)) return value.length === 0;
      return !value || (typeof value === 'object' && Object.keys(value).length === 0);
    });
    if (missing.length) {
      billingLogger_.log('[billing] Prepared payload is incomplete (missing: ' + missing.join(', ') + ')');
    }
    if (normalized && normalized.carryOverLedgerMeta && (normalized.carryOverLedgerMeta.wasAutoCreated || normalized.carryOverLedgerMeta.headerInserted)) {
      billingLogger_.log('[billing] CarryOverLedger sheet missing → using fallback model');
    }
  }

  function attachBankFlagsToBillingJson_(billingJson, bankFlagsByPatient) {
    if (!Array.isArray(billingJson)) return [];
    const flagsByPatient = bankFlagsByPatient && typeof bankFlagsByPatient === 'object' ? bankFlagsByPatient : {};
    billingJson.forEach((entry, idx) => {
      const pid = entry && entry.patientId ? String(entry.patientId).trim() : '';
      const flags = pid && flagsByPatient[pid];
      if (!flags || !entry || typeof entry !== 'object' || Array.isArray(entry)) return;
      entry.bankFlags = flags;
    });
    return billingJson;
  }

  function exportBankTransferDataForPrepared_(prepared) {
    try {
      billingLogger_.log('[billing][legacy] exportBankTransferDataForPrepared_ invoked (bank transfer export deprecated)');
    } catch (err) {
      // ignore logging errors in non-GAS environments
    }

    const normalized = normalizePreparedBilling_(prepared);
    if (!normalized) {
      billingLogger_.log('[billing] exportBankTransferDataForPrepared_: normalized payload missing', {
        hasPrepared: !!prepared,
        preparedKeys: prepared && typeof prepared === 'object' ? Object.keys(prepared) : null
      });
      throw new Error('銀行データを生成できません。請求データが未生成です。先に「請求データを集計」を実行してください。');
    }

    if (!Array.isArray(normalized.billingJson)) {
      billingLogger_.log('[billing] exportBankTransferDataForPrepared_: billingJson missing or invalid', {
        billingJsonType: typeof normalized.billingJson,
        hasBillingJson: !!normalized.billingJson
      });
      throw new Error('銀行データを生成できません。請求データの形式が不正です。先に「請求データを集計」を実行してください。');
    }
    logPreparedBankPayloadStatus_(normalized);

    if (normalized.billingJson.length === 0) {
      billingLogger_.log('[billing] exportBankTransferDataForPrepared_: billingJson empty for ' + (normalized.billingMonth || ''));
      return { billingMonth: normalized.billingMonth || '', rows: [], inserted: 0, skipped: 0, message: '当月の請求対象はありません' };
    }

    let bankInfoByName = normalized.bankInfoByName || {};
    let patientMap = normalized.patients || normalized.patientMap || {};
    let bankStatuses = normalized.bankStatuses || {};
    let bankFlagsByPatient = normalized.bankFlagsByPatient || {};

    if (!Object.keys(bankInfoByName).length || !Object.keys(patientMap).length) {
      const source = getBillingSourceData(normalized.billingMonth);
      bankInfoByName = source.bankInfoByName || bankInfoByName;
      patientMap = source.patients || source.patientMap || patientMap;
      bankStatuses = source.bankStatuses || bankStatuses;
      bankFlagsByPatient = source.bankFlagsByPatient || bankFlagsByPatient;
      if (!normalized.billingJson || !normalized.billingJson.length) {
        normalized.billingJson = generateBillingJsonFromSource(source);
      }
    }

    const billingJsonWithFlags = attachBankFlagsToBillingJson_(normalized.billingJson, bankFlagsByPatient);
    const buildResult = buildBankTransferRowsForBilling_(billingJsonWithFlags, bankInfoByName, patientMap, normalized.billingMonth, bankStatuses);
    const outputResult = exportBankTransferRows_(buildResult.billingMonth, buildResult.rows, bankStatuses);
    const combined = Object.assign({}, buildResult, outputResult);

    if (!buildResult.rows.length) {
      const reasonSummary = buildResult.skipReasons || {};
      const parts = [
        '銀行CSVが生成されませんでした',
        `総件数: ${buildResult.total || 0}`,
        `有効: ${buildResult.passed || 0}`,
        `銀行コード不正: ${reasonSummary.invalidBankCode || 0}`,
        `支店コード不正: ${reasonSummary.invalidBranchCode || 0}`,
        `口座番号不正: ${reasonSummary.invalidAccountNumber || 0}`
      ];
      combined.message = parts.join(' / ');
    }

    return combined;
  }

/***** Billing history and payment result utilities (retained for compatibility) *****/

const BILLING_HISTORY_HEADERS = [
  'billingMonth',
  'patientId',
  'nameKanji',
  'billingAmount',
  'carryOverAmount',
  'grandTotal',
  'paidAmount',
  'unpaidAmount',
  'bankStatus',
  'updatedAt',
  'memo',
  'receiptStatus',
  'aggregateUntilMonth',
  'previousReceiptAmount'
];

function resolveBillingHistoryColumns_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), BILLING_HISTORY_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const workingHeaders = headers.slice();
  const columns = {};

  BILLING_HISTORY_HEADERS.forEach((label) => {
    let idx = workingHeaders.indexOf(label);
    if (idx >= 0) {
      columns[label] = idx + 1;
      return;
    }
    const newIndex = workingHeaders.length + 1;
    sheet.getRange(1, newIndex).setValue(label);
    workingHeaders.push(label);
    columns[label] = newIndex;
  });

  return { columns, headers: workingHeaders };
}

function ensureBillingHistorySheet_() {
  const SHEET_NAME = '請求履歴';
  const workbook = resolveBillingSpreadsheetForOutput_();
  if (!workbook) return null;
  let sheet = workbook.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, BILLING_HISTORY_HEADERS.length).setValues([BILLING_HISTORY_HEADERS]);
  }
  const resolved = resolveBillingHistoryColumns_(sheet);
  return { sheet, columns: resolved.columns, headers: resolved.headers };
}

function ensureUnpaidHistorySheet_() {
  const workbook = resolveBillingSpreadsheetForOutput_();
  if (!workbook) return null;
  let sheet = workbook.getSheetByName(UNPAID_HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(UNPAID_HISTORY_SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([[
      'patientId',
      '対象月',
      '金額',
      '理由',
      '備考',
      '記録日時'
    ]]);
  }
  return sheet;
}

function appendUnpaidHistoryEntries_(entries) {
  const sheet = ensureUnpaidHistorySheet_();
  if (!sheet) {
    return { added: 0, skipped: entries && entries.length ? entries.length : 0 };
  }
  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 6).getValues()
    : [];
  const existingKeys = existing.reduce((set, row) => {
    const pid = billingNormalizePatientId_(row[0]);
    const month = String(row[1] || '').trim();
    const amount = Number(row[2]) || 0;
    if (pid && month) set[pid + '::' + month + '::' + amount] = true;
    return set;
  }, {});

  const rows = (entries || [])
    .map(entry => ({
      patientId: billingNormalizePatientId_(entry && entry.patientId),
      billingMonth: entry && entry.billingMonth ? String(entry.billingMonth).trim() : '',
      amount: entry && entry.unpaidAmount != null ? Number(entry.unpaidAmount) || 0 : 0,
      reason: entry && entry.reason ? String(entry.reason).trim() : '',
      memo: entry && entry.memo ? String(entry.memo).trim() : ''
    }))
    .filter(entry => entry.patientId && entry.billingMonth && entry.amount)
    .filter(entry => !existingKeys[entry.patientId + '::' + entry.billingMonth + '::' + entry.amount]);

  if (!rows.length) {
    return { added: 0, skipped: entries && entries.length ? entries.length : 0 };
  }

  const values = rows.map(entry => [
    entry.patientId,
    entry.billingMonth,
    entry.amount,
    entry.reason || BANK_WITHDRAWAL_UNPAID_HEADER,
    entry.memo || '',
    new Date()
  ]);

  sheet.insertRows(2, values.length);
  sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  return { added: values.length, skipped: (entries || []).length - values.length };
}

function appendBillingHistoryRows(billingJson, options) {
  const opts = options || {};
  const ensured = ensureBillingHistorySheet_();
  if (!ensured) return { billingMonth: opts.billingMonth || '', inserted: 0 };
  const sheet = ensured.sheet;
  const columns = ensured.columns;
  const headers = ensured.headers;
  const colCount = Math.max(sheet.getLastColumn(), headers.length, Math.max.apply(null, Object.values(columns)));
  const billingMonth = opts.billingMonth || (Array.isArray(billingJson) && billingJson.length && billingJson[0].billingMonth) || '';
  const memoProvided = Object.prototype.hasOwnProperty.call(opts, 'memo');
  const memo = memoProvided ? opts.memo : null;

  const lastRow = sheet.getLastRow();
  const existingValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, colCount).getValues() : [];
  const keyForRow = (row) => {
    const month = columns.billingMonth ? String(row[columns.billingMonth - 1] || '').trim() : '';
    const pid = columns.patientId ? String(row[columns.patientId - 1] || '').trim() : '';
    return month && pid ? `${month}::${pid}` : '';
  };

  const workingRowsByKey = existingValues.reduce((map, row) => {
    const key = keyForRow(row);
    if (key && !map.has(key)) map.set(key, row);
    return map;
  }, new Map());

  const mapped = (billingJson || []).map(item => {
    const billingAmount = item && item.treatmentAmount != null ? item.treatmentAmount : 0;
    const carryOver = item && item.carryOverAmount != null ? item.carryOverAmount : 0;
    const paid = item && item.paidAmount != null ? item.paidAmount : 0;
    const grandTotal = normalizeBillingAmount_(item);
    const unpaid = grandTotal - paid;
    const row = new Array(colCount).fill('');
    if (columns.billingMonth) row[columns.billingMonth - 1] = billingMonth;
    if (columns.patientId) row[columns.patientId - 1] = item && item.patientId ? item.patientId : '';
    if (columns.nameKanji) row[columns.nameKanji - 1] = item && item.nameKanji ? item.nameKanji : '';
    if (columns.billingAmount) row[columns.billingAmount - 1] = billingAmount;
    if (columns.carryOverAmount) row[columns.carryOverAmount - 1] = carryOver;
    if (columns.grandTotal) row[columns.grandTotal - 1] = grandTotal;
    if (columns.paidAmount) row[columns.paidAmount - 1] = paid;
    if (columns.unpaidAmount) row[columns.unpaidAmount - 1] = unpaid;
    if (columns.bankStatus) row[columns.bankStatus - 1] = item && item.bankStatus ? item.bankStatus : '';
    if (columns.updatedAt) row[columns.updatedAt - 1] = new Date();
    if (columns.memo && memoProvided) row[columns.memo - 1] = memo || '';
    if (columns.receiptStatus) row[columns.receiptStatus - 1] = item && item.receiptStatus ? item.receiptStatus : '';
    if (columns.aggregateUntilMonth) row[columns.aggregateUntilMonth - 1] = item && item.aggregateUntilMonth ? item.aggregateUntilMonth : '';

    return { key: (billingMonth && item && item.patientId) ? `${billingMonth}::${item.patientId}` : keyForRow(row), row };
  });

  mapped.forEach(entry => {
    const existingRow = workingRowsByKey.get(entry.key);
    const mergedRow = existingRow ? existingRow.slice() : new Array(colCount).fill('');

    const applyValue = (label, value) => {
      const idx = columns[label];
      if (!idx) return;
      mergedRow[idx - 1] = value;
    };

    applyValue('billingMonth', entry.row[columns.billingMonth - 1]);
    applyValue('patientId', entry.row[columns.patientId - 1]);
    applyValue('nameKanji', entry.row[columns.nameKanji - 1]);
    applyValue('billingAmount', entry.row[columns.billingAmount - 1]);
    applyValue('carryOverAmount', entry.row[columns.carryOverAmount - 1]);
    applyValue('grandTotal', entry.row[columns.grandTotal - 1]);
    applyValue('paidAmount', entry.row[columns.paidAmount - 1]);
    applyValue('unpaidAmount', entry.row[columns.unpaidAmount - 1]);
    applyValue('bankStatus', entry.row[columns.bankStatus - 1]);
    applyValue('updatedAt', entry.row[columns.updatedAt - 1]);
    applyValue('previousReceiptAmount', entry.row[columns.previousReceiptAmount - 1]);

    if (columns.memo) {
      const existingMemo = existingRow ? existingRow[columns.memo - 1] : '';
      const resolvedMemo = memoProvided ? (memo || '') : existingMemo;
      mergedRow[columns.memo - 1] = resolvedMemo;
    }

    if (columns.receiptStatus) {
      mergedRow[columns.receiptStatus - 1] = entry.row[columns.receiptStatus - 1];
    }

    if (columns.aggregateUntilMonth) {
      mergedRow[columns.aggregateUntilMonth - 1] = entry.row[columns.aggregateUntilMonth - 1];
    }

    workingRowsByKey.set(entry.key, mergedRow);
  });

  const workingRows = Array.from(workingRowsByKey.values());
  const dataRowCount = Math.max(0, lastRow - 1);
  const maxRowCount = Math.max(dataRowCount, workingRows.length);

  if (maxRowCount > 0) {
    sheet.getRange(2, 1, maxRowCount, colCount).clearContent();
  }
  if (workingRows.length) {
    sheet.getRange(2, 1, workingRows.length, colCount).setValues(workingRows);
  }

  return { billingMonth, inserted: mapped.length };
}

function applyPaymentResultsToHistory(billingMonth, bankStatuses) {
  const ensured = ensureBillingHistorySheet_();
  if (!ensured) return { billingMonth, updated: 0 };
  const sheet = ensured.sheet;
  const columns = ensured.columns;
  const headers = ensured.headers;
  const colCount = Math.max(sheet.getLastColumn(), headers.length, Math.max.apply(null, Object.values(columns)));
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { billingMonth, updated: 0 };
  const data = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
  const updates = [];
  const statusMap = bankStatuses || {};
  data.forEach((row, idx) => {
    if (columns.billingMonth && row[columns.billingMonth - 1] !== billingMonth) return;
    const pid = columns.patientId ? row[columns.patientId - 1] : '';
    const statusEntry = statusMap[pid];
    if (!statusEntry) return;
    const newRow = row.slice();
    let changed = false;

    if (columns.bankStatus && statusEntry.bankStatus) {
      newRow[columns.bankStatus - 1] = statusEntry.bankStatus;
      changed = true;
    }
    if (columns.paidAmount && statusEntry.paidAmount != null) {
      const paid = Number(statusEntry.paidAmount) || 0;
      newRow[columns.paidAmount - 1] = paid;
      const grandTotal = columns.grandTotal ? Number(newRow[columns.grandTotal - 1]) || 0 : 0;
      const unpaid = statusEntry.unpaidAmount != null ? statusEntry.unpaidAmount : grandTotal - paid;
      if (columns.unpaidAmount) newRow[columns.unpaidAmount - 1] = unpaid;
      changed = true;
    }

    if (!changed) return;
    if (columns.updatedAt) newRow[columns.updatedAt - 1] = new Date();
    updates.push({ rowNumber: idx + 2, values: newRow });
  });
  updates.forEach(update => {
    sheet.getRange(update.rowNumber, 1, 1, colCount).setValues([update.values]);
  });
  return { billingMonth, updated: updates.length };
}

const BILLING_PAYMENT_PDF_STATUS_LABELS = {
  '回収済み': 'OK',
  '預金口座振替依頼書なし': 'NO_DOCUMENT',
  '資金不足': 'INSUFFICIENT',
  '取引なし': 'NOT_FOUND'
};

function parseBillingPaymentResultPdf(pdfBlob) {
  const content = pdfBlob.getDataAsString();
  const lines = content.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  const entries = [];
  lines.forEach(line => {
    const paidMatch = line.match(/([0-9,]+)円/);
    const statusLabel = Object.keys(BILLING_PAYMENT_PDF_STATUS_LABELS).find(label => line.indexOf(label) >= 0) || '';
    const nameMatch = line.match(/^(.*?)(回収済み|預金口座振替依頼書なし|資金不足|取引なし)/);
    if (!nameMatch) return;
    const name = normalizeBillingNameKey_(nameMatch[1]);
    const paidAmount = paidMatch ? Number(paidMatch[1].replace(/,/g, '')) : 0;
    const bankStatus = BILLING_PAYMENT_PDF_STATUS_LABELS[statusLabel] || '';
    entries.push({ nameKanji: name, paidAmount, bankStatus, statusLabel });
  });
  return entries;
}

function applyPaymentResultPdf(billingMonth, pdfBlob, billingJson) {
  const parsed = parseBillingPaymentResultPdf(pdfBlob);
  const nameIndex = {};
  (billingJson || []).forEach(item => {
    const key = normalizeBillingNameKey_(item.nameKanji);
    if (key && !nameIndex[key]) {
      nameIndex[key] = item;
    }
  });

  const matched = [];
  parsed.forEach(entry => {
    const key = normalizeBillingNameKey_(entry.nameKanji);
    const target = nameIndex[key];
    if (!target) return;
    matched.push({
      patientId: target.patientId,
      billingMonth,
      paidAmount: entry.paidAmount,
      unpaidAmount: normalizeBillingAmount_(target) - entry.paidAmount,
      bankStatus: entry.bankStatus,
      statusLabel: entry.statusLabel
    });
  });

  const statusMap = matched.reduce((map, entry) => {
    map[entry.patientId] = entry;
    return map;
  }, {});
  const historyResult = applyPaymentResultsToHistory(billingMonth, statusMap);

  return {
    billingMonth,
    parsedCount: parsed.length,
    matched: matched.length,
    updated: historyResult.updated,
    entries: matched
  };
}
