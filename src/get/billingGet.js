/***** Get layer: billing data retrieval *****/

const BILLING_PATIENT_RAW_COL_LIMIT = columnLetterToNumber_('BJ');
const BILLING_TREATMENT_SHEET_NAME = '施術録';
const BILLING_PAYMENT_RESULT_SHEET_PREFIX = '入金結果_';
const BILLING_PATIENT_SHEET_NAME = '患者情報';
const BILLING_BANK_SHEET_NAME = '銀行情報';
const BILLING_BANK_STATUS_ALLOWLIST = ['OK', 'NO_DOCUMENT', 'INSUFFICIENT', 'NOT_FOUND'];

/**
 * YYYYMM 形式の請求月を正規化
 */
function normalizeBillingMonthInput(billingMonth) {
  if (billingMonth && typeof billingMonth === 'object' && billingMonth.key && billingMonth.start && billingMonth.end) {
    return billingMonth;
  }
  if (!billingMonth) {
    throw new Error('請求月が指定されていません');
  }
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  let year = null;
  let month = null;

  if (billingMonth instanceof Date && !isNaN(billingMonth.getTime())) {
    year = billingMonth.getFullYear();
    month = billingMonth.getMonth() + 1;
  } else {
    const raw = String(billingMonth).trim();
    const normalizedDigits = raw.replace(/\D/g, '');
    if (normalizedDigits.length === 6) {
      year = Number(normalizedDigits.slice(0, 4));
      month = Number(normalizedDigits.slice(4, 6));
    } else {
      const match = raw.match(/^(\d{4})\s*[\/-]?\s*(\d{1,2})$/);
      if (match) {
        year = Number(match[1]);
        month = Number(match[2]);
      }
    }
  }

  if (!year || !month || month < 1 || month > 12) {
    throw new Error('請求月の形式が不正です: ' + billingMonth);
  }

  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 1);
  const key = Utilities.formatDate(start, tz, 'yyyyMM');

  return { year, month, key, start, end, timezone: tz };
}

function columnLetterToNumber_(letter) {
  const raw = String(letter || '').trim().toUpperCase();
  if (!raw || !/^[A-Z]+$/.test(raw)) return null;
  let num = 0;
  for (let i = 0; i < raw.length; i++) {
    num = num * 26 + (raw.charCodeAt(i) - 64);
  }
  return num;
}

function columnNumberToLetter_(num) {
  const n = Number(num);
  if (!isFinite(n) || n <= 0) return '';
  let x = Math.floor(n);
  let letters = '';
  while (x > 0) {
    const rem = (x - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    x = Math.floor((x - 1) / 26);
  }
  return letters;
}

function resolveBillingColumn_(headers, labelCandidates, fieldLabel, options) {
  const opts = options || {};
  const headerMap = buildHeaderMap_(headers);
  for (let i = 0; i < labelCandidates.length; i++) {
    const key = normalizeHeaderKey_(labelCandidates[i]);
    if (key && headerMap[key]) {
      return headerMap[key];
    }
  }
  if (opts.fallbackIndex && headers.length >= opts.fallbackIndex) {
    return opts.fallbackIndex;
  }
  if (opts.fallbackLetter) {
    const idx = columnLetterToNumber_(opts.fallbackLetter);
    if (idx && headers.length >= idx) {
      return idx;
    }
  }
  if (opts.required) {
    throw new Error(fieldLabel + '列が見つかりません');
  }
  return null;
}

function buildPatientRawObject_(headers, rowValues) {
  const raw = {};
  const len = Math.min(headers.length, rowValues.length);
  for (let i = 0; i < len; i++) {
    const letterKey = columnNumberToLetter_(i + 1);
    raw[letterKey] = rowValues[i];
    const headerText = headers[i] ? String(headers[i]).trim() : '';
    const normalizedHeader = headerText || letterKey;
    if (!raw.hasOwnProperty(normalizedHeader)) {
      raw[normalizedHeader] = rowValues[i];
    }
  }
  return raw;
}

function normalizeBillingNameKey_(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}

function normalizeDisabledFlag_(value) {
  const raw = value == null ? '' : String(value).trim();
  if (!raw) return 0;
  const num = Number(raw);
  if (Number.isFinite(num)) {
    return num;
  }
  if (['無効', '停止', '中止', 'disabled'].indexOf(raw.toLowerCase()) >= 0) {
    return 2;
  }
  return 0;
}

function normalizeBurdenRateInt_(value) {
  if (value == null || value === '') return 0;
  if (typeof value === 'number') {
    if (!isFinite(value)) return 0;
    if (value <= 1) return Math.round(value * 10);
    if (value <= 3) return Math.round(value);
    if (value === 10 || value === 20 || value === 30) return value / 10;
    if (value === 0) return 0;
  }
  const text = String(value).trim();
  if (!text) return 0;
  const digits = text.replace(/[^0-9.]/g, '');
  if (digits) {
    const num = Number(digits);
    if (!isNaN(num)) {
      if (num === 0) return 0;
      if (num <= 1) return Math.round(num * 10);
      if (num <= 3) return Math.round(num);
      if (num === 10 || num === 20 || num === 30) return num / 10;
    }
  }
  const ratio = normalizeBurdenRatio_(text);
  if (ratio === 0) return 0;
  if (ratio != null) return Math.round(ratio * 10);
  return 0;
}

function normalizeMoneyValue_(value) {
  if (typeof value === 'number') {
    return isFinite(value) ? value : 0;
  }
  const text = String(value || '').replace(/,/g, '').trim();
  if (!text) return 0;
  const num = Number(text);
  return isNaN(num) ? 0 : num;
}

function normalizeZeroOneFlag_(value) {
  if (value === true) return 1;
  if (value === false) return 0;
  if (typeof value === 'number') return value ? 1 : 0;
  const text = String(value || '').trim().toLowerCase();
  if (!text) return 0;
  if (['1', 'true', 'yes', 'y', 'on', 'new', '新規', '〇', '○', '◯'].includes(text)) return 1;
  if (['0', 'false', 'no', 'off', '旧', '既存'].includes(text)) return 0;
  const num = Number(text);
  return isNaN(num) ? 0 : (num ? 1 : 0);
}

function normalizeBankStatus_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const normalized = raw.toUpperCase().replace(/[-\s]+/g, '_');
  if (BILLING_BANK_STATUS_ALLOWLIST.indexOf(normalized) >= 0) {
    return normalized;
  }
  return normalized;
}

function indexByPatientId_(records) {
  return records.reduce((map, record) => {
    if (record && record.patientId) {
      map[record.patientId] = record;
    }
    return map;
  }, {});
}

function getBillingTreatmentVisitCounts(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const sheet = ss().getSheetByName(BILLING_TREATMENT_SHEET_NAME);
  if (!sheet) {
    throw new Error('施術録シートが見つかりません: ' + BILLING_TREATMENT_SHEET_NAME);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const width = Math.min(Math.max(sheet.getLastColumn(), 2), sheet.getMaxColumns());
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  const colDate = resolveBillingColumn_(headers, ['タイムスタンプ', '日付', '施術日', '記録日', '日時'], '日付', {
    required: true,
    fallbackIndex: 1
  });
  const colPid = resolveBillingColumn_(headers, LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', {
    required: true,
    fallbackIndex: 2
  });

  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const counts = {};
  values.forEach(row => {
    const pid = normId_(row[colPid - 1]);
    if (!pid) return;
    const dateCell = row[colDate - 1];
    const d = dateCell instanceof Date ? dateCell : parseDateFlexible_(dateCell);
    if (!(d instanceof Date) || isNaN(d.getTime())) return;
    if (d < month.start || d >= month.end) return;
    const current = counts[pid] || { visitCount: 0 };
    current.visitCount += 1;
    counts[pid] = current;
  });
  return counts;
}

function getBillingPatientRecords() {
  const sheet = sh(BILLING_PATIENT_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rawColCount = Math.min(sheet.getLastColumn(), BILLING_PATIENT_RAW_COL_LIMIT);
  const headers = sheet.getRange(1, 1, 1, rawColCount).getDisplayValues()[0];
  const values = sheet.getRange(2, 1, lastRow - 1, rawColCount).getValues();

  const colPid = resolveBillingColumn_(headers, LABELS.recNo, '患者ID', { required: true, fallbackIndex: PATIENT_COLS_FIXED.recNo });
  const colName = resolveBillingColumn_(headers, LABELS.name, '名前', { fallbackIndex: PATIENT_COLS_FIXED.name });
  const colKana = resolveBillingColumn_(headers, LABELS.furigana, 'フリガナ', { fallbackIndex: PATIENT_COLS_FIXED.furigana });
  const colInsurance = resolveBillingColumn_(headers, ['保険区分', '保険種別', '保険タイプ', '保険'], '保険区分', {});
  const colBurden = resolveBillingColumn_(headers, LABELS.share, '負担割合', { fallbackIndex: PATIENT_COLS_FIXED.share });
  const colUnitPrice = resolveBillingColumn_(headers, ['単価', '請求単価', '自費単価', '単価(自費)', '単価（自費）'], '単価', {});
  const colAddress = resolveBillingColumn_(headers, ['住所', '住所1', '住所２', '住所2', 'address', 'Address'], '住所', {});
  const colBank = resolveBillingColumn_(headers, ['銀行コード', '銀行CD', '銀行番号', 'bankCode'], '銀行コード', { fallbackLetter: 'N' });
  const colBranch = resolveBillingColumn_(headers, ['支店コード', '支店番号', '支店CD', 'branchCode'], '支店コード', { fallbackLetter: 'O' });
  const colAccount = resolveBillingColumn_(headers, ['口座番号', '口座No', '口座NO', 'accountNumber', '口座'], '口座番号', { fallbackLetter: 'Q' });
  const colIsNew = resolveBillingColumn_(headers, ['新規', '新患', 'isNew', '新規フラグ', '新規区分'], '新規区分', { fallbackLetter: 'U' });
  const colCarryOver = resolveBillingColumn_(headers, ['未入金', '未入金額', '未収金', '未収', '繰越', '繰越額', '繰り越し', '差引繰越', '前回未払', '前回未収', 'carryOverAmount'], '未入金額', {});

  return values.map(row => {
    const pid = normId_(row[colPid - 1]);
    if (!pid) return null;
    return {
      patientId: pid,
      raw: buildPatientRawObject_(headers, row),
      nameKanji: colName ? String(row[colName - 1] || '').trim() : '',
      nameKana: colKana ? String(row[colKana - 1] || '').trim() : '',
      insuranceType: colInsurance ? String(row[colInsurance - 1] || '').trim() : '',
      burdenRate: colBurden ? normalizeBurdenRateInt_(row[colBurden - 1]) : 0,
      unitPrice: colUnitPrice ? normalizeMoneyValue_(row[colUnitPrice - 1]) : 0,
      address: colAddress ? String(row[colAddress - 1] || '').trim() : '',
      bankCode: colBank ? String(row[colBank - 1] || '').trim() : '',
      branchCode: colBranch ? String(row[colBranch - 1] || '').trim() : '',
      accountNumber: colAccount ? String(row[colAccount - 1] || '').trim() : '',
      isNew: colIsNew ? normalizeZeroOneFlag_(row[colIsNew - 1]) : 0,
      carryOverAmount: colCarryOver ? normalizeMoneyValue_(row[colCarryOver - 1]) : 0
    };
  }).filter(Boolean);
}

function getBillingBankRecords() {
  const sheet = ss().getSheetByName(BILLING_BANK_SHEET_NAME);
  if (!sheet) {
    throw new Error('銀行情報シートが見つかりません: ' + BILLING_BANK_SHEET_NAME);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  const colName = resolveBillingColumn_(headers, LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const colKana = resolveBillingColumn_(headers, LABELS.furigana, 'フリガナ', { fallbackLetter: 'B' });
  const colKanaAlt = resolveBillingColumn_(headers, LABELS.furigana, 'フリガナ', { fallbackLetter: 'R' });
  const colBank = resolveBillingColumn_(headers, ['銀行コード', '銀行CD', '銀行番号', 'bankCode'], '銀行コード', { required: true, fallbackLetter: 'N' });
  const colBranch = resolveBillingColumn_(headers, ['支店コード', '支店番号', '支店CD', 'branchCode'], '支店コード', { required: true, fallbackLetter: 'O' });
  const colRegulation = resolveBillingColumn_(headers, ['規定コード', '規定', '規定CD', '規定コード(1固定)', '1固定'], '規定コード', { fallbackLetter: 'P' });
  const colAccount = resolveBillingColumn_(headers, ['口座番号', '口座No', '口座NO', 'accountNumber', '口座'], '口座番号', { required: true, fallbackLetter: 'Q' });
  const colIsNew = resolveBillingColumn_(headers, ['新規', '新患', 'isNew', '新規フラグ', '新規区分'], '新規区分', { fallbackLetter: 'U' });
  const colDisabled = resolveBillingColumn_(headers, ['利用停止', '停止', '無効', 'ステータス'], '利用停止', { fallbackLetter: 'T' });

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.map(row => {
    const nameKanji = colName ? String(row[colName - 1] || '').trim() : '';
    if (!nameKanji) return null;
    const disabledFlag = colDisabled ? normalizeDisabledFlag_(row[colDisabled - 1]) : 0;
    if (disabledFlag === 2) return null;
    const kanaPrimary = colKana ? String(row[colKana - 1] || '').trim() : '';
    const kanaSecondary = colKanaAlt ? String(row[colKanaAlt - 1] || '').trim() : '';
    const regulationCode = colRegulation ? normalizeMoneyValue_(row[colRegulation - 1]) : 0;
    const isNew = colIsNew ? normalizeZeroOneFlag_(row[colIsNew - 1]) : 0;
    return {
      nameKanji,
      nameKana: kanaPrimary || kanaSecondary,
      bankCode: colBank ? String(row[colBank - 1] || '').trim() : '',
      branchCode: colBranch ? String(row[colBranch - 1] || '').trim() : '',
      regulationCode: regulationCode || 1,
      accountNumber: colAccount ? String(row[colAccount - 1] || '').trim() : '',
      isNew,
      raw: buildPatientRawObject_(headers, row)
    };
  }).filter(Boolean);
}

function buildBankLookupByKanji_(bankRecords) {
  return (bankRecords || []).reduce((map, rec) => {
    if (!rec) return map;
    const key = normalizeBillingNameKey_(rec.nameKanji);
    if (key && !map[key]) {
      map[key] = rec;
    }
    return map;
  }, {});
}

function getBillingPaymentResults(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const sheetName = BILLING_PAYMENT_RESULT_SHEET_PREFIX + month.key;
  const sheet = ss().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('入金結果シートが見つかりません: ' + sheetName);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const colPid = resolveBillingColumn_(headers, LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', { required: true });
  const colStatus = resolveBillingColumn_(headers, ['bankStatus', '入金ステータス', 'ステータス', '状態', '結果'], '入金ステータス', { required: true });
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map = {};
  values.forEach(row => {
    const pid = normId_(row[colPid - 1]);
    const status = normalizeBankStatus_(row[colStatus - 1]);
    if (!pid || !status) return;
    map[pid] = { bankStatus: status };
  });
  return map;
}

function getBillingSourceData(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const patientRecords = getBillingPatientRecords();
  const bankRecords = getBillingBankRecords();
  const patientMap = indexByPatientId_(patientRecords);
  return {
    billingMonth: month.key,
    month,
    treatmentVisitCounts: getBillingTreatmentVisitCounts(month),
    patients: patientMap,
    bankInfoByName: buildBankLookupByKanji_(bankRecords),
    bankStatuses: getBillingPaymentResults(month)
  };
}
