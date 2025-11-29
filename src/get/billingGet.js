/***** Get layer: billing data retrieval *****/

/**
 * Provide local fallbacks for shared helpers so the billing pipeline can run
 * without depending on Code.js ordering.
 */
const billingNormalizeHeaderKey_ = typeof normalizeHeaderKey_ === 'function'
  ? normalizeHeaderKey_
  : function normalizeHeaderKey_(s) {
    if (!s) return '';
    const z2h = String(s).normalize('NFKC');
    const noSpace = z2h.replace(/\s+/g, '');
    const noPunct = noSpace.replace(/[（）\(\)\[\]【】:：・\-＿_]/g, '');
    return noPunct.toLowerCase();
  };

const billingNormalizePatientId_ = typeof normId_ === 'function'
  ? normId_
  : function normIdFallback_(value) {
    return String(value || '').trim();
  };

const billingParseDateFlexible_ = typeof parseDateFlexible_ === 'function'
  ? parseDateFlexible_
  : function parseDateFlexibleFallback_(value) {
    if (value instanceof Date) return value;
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
  };

const billingBuildHeaderMap_ = typeof buildHeaderMap_ === 'function'
  ? buildHeaderMap_
  : function buildHeaderMap_(headersRow) {
    const map = {};
    (headersRow || []).forEach((header, idx) => {
      const key = billingNormalizeHeaderKey_(header);
      if (key && !map[key]) map[key] = idx + 1;
    });
    return map;
  };

const billingNormalizeVisitCount_ = typeof normalizeVisitCount_ === 'function'
  ? normalizeVisitCount_
  : function normalizeVisitCount_(value) {
    const num = Number(value && value.visitCount != null ? value.visitCount : value);
    return Number.isFinite(num) && num > 0 ? num : 0;
  };

const billingNormalizeBurdenRatio_ = typeof normalizeBurdenRatio_ === 'function'
  ? normalizeBurdenRatio_
  : function normalizeBurdenRatio_(text) {
    if (!text) return null;
    const normalized = String(text).normalize('NFKC').replace(/\s/g, '').replace('％', '%').replace('割', '');
    if (/^[123]$/.test(normalized)) return Number(normalized) / 10;
    if (/^(10|20|30)%?$/.test(normalized)) return Number(RegExp.$1) / 100;
    return null;
  };

const BILLING_LABELS = typeof LABELS !== 'undefined' ? LABELS : {
  recNo: ['施術録番号', '施術録No', '施術録NO', '記録番号', 'カルテ番号', '患者ID', '患者番号'],
  name: ['名前', '氏名', '患者名', 'お名前'],
  hospital: ['病院名', '医療機関', '病院'],
  doctor: ['医師', '主治医', '担当医'],
  furigana: ['ﾌﾘｶﾞﾅ', 'ふりがな', 'フリガナ'],
  birth: ['生年月日', '誕生日', '生年', '生年月'],
  consent: ['同意年月日', '同意日', '同意開始日', '同意開始'],
  consentHandout: ['配布', '配布欄', '配布状況', '配布日', '配布（同意書）'],
  consentContent: ['同意症状', '同意内容', '施術対象疾患', '対象疾患', '対象症状', '同意書内容', '同意記載内容'],
  share: ['負担割合', '負担', '自己負担', '負担率', '負担割', '負担%', '負担％'],
  phone: ['電話', '電話番号', 'TEL', 'Tel']
};

const BILLING_PATIENT_COLS_FIXED = typeof PATIENT_COLS_FIXED !== 'undefined' ? PATIENT_COLS_FIXED : {
  recNo: 3,
  name: 4,
  hospital: 5,
  furigana: 6,
  birth: 7,
  doctor: 26,
  consent: 28,
  consentHandout: 54,
  consentContent: 25,
  phone: 32,
  share: 47
};

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

const billingSs = resolveBillingSpreadsheet_;

const BILLING_PATIENT_RAW_COL_LIMIT = columnLetterToNumber_('BJ');
const BILLING_TREATMENT_SHEET_NAME = '施術録';
const BILLING_PAYMENT_RESULT_SHEET_PREFIX = '入金結果_';
const BILLING_PATIENT_SHEET_NAME = '患者情報';
const BILLING_BANK_SHEET_NAME = '銀行情報';
const BILLING_BANK_STATUS_ALLOWLIST = ['OK', 'NO_DOCUMENT', 'INSUFFICIENT', 'NOT_FOUND'];
const BILLING_PAID_STATUS_ALLOWLIST = ['回収', '未回収', '手続き中', '手続中', 'エラー'];

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
  const headerMap = billingBuildHeaderMap_(headers);
  for (let i = 0; i < labelCandidates.length; i++) {
    const key = billingNormalizeHeaderKey_(labelCandidates[i]);
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
  const text = String(value).normalize('NFKC').trim();
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
  const ratio = billingNormalizeBurdenRatio_(text);
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

function normalizePaidStatus_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const normalized = raw.replace(/\s+/g, '');
  const matched = BILLING_PAID_STATUS_ALLOWLIST.find(label => label === normalized);
  if (matched) {
    return matched === '手続中' ? '手続き中' : matched;
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

function loadTreatmentLogs_() {
  const sheet = billingSs().getSheetByName(BILLING_TREATMENT_SHEET_NAME);
  if (!sheet) {
    throw new Error('施術録シートが見つかりません: ' + BILLING_TREATMENT_SHEET_NAME);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const width = Math.min(Math.max(sheet.getLastColumn(), 2), sheet.getMaxColumns());
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  const colDate = resolveBillingColumn_(headers, ['タイムスタンプ', '日付', '施術日', '記録日', '日時'], '日付', {
    required: true,
    fallbackIndex: 1
  });
  const colPid = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', {
    required: true,
    fallbackIndex: 2
  });
  const colCreatedBy = resolveBillingColumn_(headers,
    ['createdbyemail', 'createdby', '記録者', '作成者', '担当者', '担当メール', '入力者', '編集者'],
    '作成者',
    { fallbackLetter: 'E' }
  );

  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  return values.map((row, idx) => {
    const pid = billingNormalizePatientId_(row[colPid - 1]);
    const dateCell = row[colDate - 1];
    const timestamp = dateCell instanceof Date ? dateCell : billingParseDateFlexible_(dateCell);
    return {
      rowNumber: idx + 2,
      patientId: pid,
      timestamp,
      createdByEmail: colCreatedBy ? String(row[colCreatedBy - 1] || '').trim() : '',
      raw: row
    };
  });
}

function buildVisitCountMap_(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const logs = loadTreatmentLogs_();
  const counts = {};
  const latestStaffByPatient = {};
  logs.forEach(log => {
    const pid = log && log.patientId ? billingNormalizePatientId_(log.patientId) : '';
    const ts = log && log.timestamp;
    if (!pid || !(ts instanceof Date) || isNaN(ts.getTime())) return;
    if (ts < month.start || ts >= month.end) return;
    const current = counts[pid] || { visitCount: 0 };
    current.visitCount += 1;
    counts[pid] = current;

    if (log && log.createdByEmail) {
      const existing = latestStaffByPatient[pid];
      if (!existing || !existing.timestamp || ts > existing.timestamp) {
        latestStaffByPatient[pid] = { email: log.createdByEmail, timestamp: ts };
      }
    }
  });
  const staffByPatient = Object.keys(latestStaffByPatient).reduce((map, pid) => {
    map[pid] = latestStaffByPatient[pid].email || '';
    return map;
  }, {});
  return { billingMonth: month.key, counts, staffByPatient };
}

function getBillingTreatmentVisitCounts(billingMonth) {
  const result = buildVisitCountMap_(billingMonth);
  return result.counts;
}

function extractMonthlyVisitCounts(billingMonth) {
  const counts = getBillingTreatmentVisitCounts(billingMonth) || {};
  const normalized = {};
  Object.keys(counts).forEach(pid => {
    const visitCount = billingNormalizeVisitCount_(counts[pid]);
    if (visitCount > 0) {
      normalized[pid] = visitCount;
    }
  });
  return normalized;
}

function getBillingPatientRecords() {
  const sheet = billingSs().getSheetByName(BILLING_PATIENT_SHEET_NAME);
  if (!sheet) {
    throw new Error('患者情報シートが見つかりません: ' + BILLING_PATIENT_SHEET_NAME);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rawColCount = Math.min(sheet.getLastColumn(), BILLING_PATIENT_RAW_COL_LIMIT);
  const headers = sheet.getRange(1, 1, 1, rawColCount).getDisplayValues()[0];
  const values = sheet.getRange(2, 1, lastRow - 1, rawColCount).getValues();

  const colPid = resolveBillingColumn_(headers, BILLING_LABELS.recNo, '患者ID', { required: true, fallbackIndex: BILLING_PATIENT_COLS_FIXED.recNo });
  const colName = resolveBillingColumn_(headers, BILLING_LABELS.name, '名前', { fallbackIndex: BILLING_PATIENT_COLS_FIXED.name });
  const colKana = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', { fallbackIndex: BILLING_PATIENT_COLS_FIXED.furigana });
  const colInsurance = resolveBillingColumn_(headers, ['保険区分', '保険種別', '保険タイプ', '保険'], '保険区分', {});
  const colBurden = resolveBillingColumn_(headers, BILLING_LABELS.share, '負担割合', { fallbackIndex: BILLING_PATIENT_COLS_FIXED.share });
  const colUnitPrice = resolveBillingColumn_(headers, ['単価', '請求単価', '自費単価', '単価(自費)', '単価（自費）'], '単価', {});
  const colAddress = resolveBillingColumn_(headers, ['住所', '住所1', '住所２', '住所2', 'address', 'Address'], '住所', {});
  const colBank = resolveBillingColumn_(headers, ['銀行コード', '銀行CD', '銀行番号', 'bankCode'], '銀行コード', { fallbackLetter: 'N' });
  const colBranch = resolveBillingColumn_(headers, ['支店コード', '支店番号', '支店CD', 'branchCode'], '支店コード', { fallbackLetter: 'O' });
  const colAccount = resolveBillingColumn_(headers, ['口座番号', '口座No', '口座NO', 'accountNumber', '口座'], '口座番号', { fallbackLetter: 'Q' });
  const colIsNew = resolveBillingColumn_(headers, ['新規', '新患', 'isNew', '新規フラグ', '新規区分'], '新規区分', { fallbackLetter: 'U' });
  const colCarryOver = resolveBillingColumn_(headers, ['未入金', '未入金額', '未収金', '未収', '繰越', '繰越額', '繰り越し', '差引繰越', '前回未払', '前回未収', 'carryOverAmount'], '未入金額', {});

  return values.map(row => {
    const pid = billingNormalizePatientId_(row[colPid - 1]);
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
  const sheet = billingSs().getSheetByName(BILLING_BANK_SHEET_NAME);
  if (!sheet) {
    throw new Error('銀行情報シートが見つかりません: ' + BILLING_BANK_SHEET_NAME);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  const colName = resolveBillingColumn_(headers, BILLING_LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const colKana = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', { fallbackLetter: 'B' });
  const colKanaAlt = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', { fallbackLetter: 'R' });
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
  const sheet = billingSs().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('入金結果シートが見つかりません: ' + sheetName);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const colPid = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', { required: true });
  const colStatus = resolveBillingColumn_(headers, ['bankStatus', '入金ステータス', 'ステータス', '状態', '結果'], '入金ステータス', { required: true });
  const colPaidStatus = resolveBillingColumn_(headers, ['領収状態', '領収', 'paidStatus'], '領収状態', {});
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map = {};
  values.forEach(row => {
    const pid = billingNormalizePatientId_(row[colPid - 1]);
    const status = normalizeBankStatus_(row[colStatus - 1]);
    if (!pid || !status) return;
    map[pid] = {
      bankStatus: status,
      paidStatus: colPaidStatus ? normalizePaidStatus_(row[colPaidStatus - 1]) : ''
    };
  });
  return map;
}

function getBillingPaymentResultsIfExists_(billingMonth) {
  try {
    return getBillingPaymentResults(billingMonth);
  } catch (e) {
    const message = e && e.message ? String(e.message) : '';
    if (message.indexOf('入金結果シートが見つかりません') >= 0) {
      return {};
    }
    throw e;
  }
}

function getBillingSourceData(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const patientRecords = getBillingPatientRecords();
  const bankRecords = getBillingBankRecords();
  const patientMap = indexByPatientId_(patientRecords);
  const visitCountsResult = buildVisitCountMap_(month);
  const treatmentVisitCounts = visitCountsResult.counts;
  return {
    billingMonth: month.key,
    month,
    treatmentVisitCounts,
    visitCounts: treatmentVisitCounts,
    patients: patientMap,
    patientMap,
    bankInfoByName: buildBankLookupByKanji_(bankRecords),
    bankStatuses: getBillingPaymentResultsIfExists_(month),
    staffByPatient: visitCountsResult.staffByPatient || {}
  };
}

function extractUnpaidBillingHistory(targetBillingMonth) {
  const month = normalizeBillingMonthInput(targetBillingMonth);
  const sheet = billingSs().getSheetByName('請求履歴');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const targetKeyNum = Number(month.key);

  return values.map(row => {
    const billingMonth = row[0] ? String(row[0]).trim() : '';
    const entryMonthNum = Number(billingMonth.replace(/\D/g, '')) || 0;
    return {
      billingMonth,
      patientId: billingNormalizePatientId_(row[1]),
      nameKanji: row[2] || '',
      billingAmount: Number(row[3]) || 0,
      carryOverAmount: Number(row[4]) || 0,
      grandTotal: Number(row[5]) || 0,
      paidAmount: Number(row[6]) || 0,
      unpaidAmount: Number(row[7]) || 0,
      bankStatus: row[8] || '',
      updatedAt: billingParseDateFlexible_(row[9]) || null,
      memo: row[10] || '',
      entryMonthNum
    };
  }).filter(entry => {
    if (!entry.patientId || !entry.billingMonth) return false;
    if (!entry.entryMonthNum || entry.entryMonthNum >= targetKeyNum) return false;
    return entry.unpaidAmount > 0;
  }).map(entry => {
    const clone = Object.assign({}, entry);
    delete clone.entryMonthNum;
    return clone;
  });
}
