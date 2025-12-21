/***** Get layer: billing data retrieval *****/

/**
 * Provide local fallbacks for shared helpers so the billing pipeline can run
 * without depending on Code.js ordering.
 */
if (typeof billingLogger_ === 'undefined') {
  const billingFallbackLog_ = typeof console !== 'undefined' && console && typeof console.log === 'function'
    ? (...args) => console.log(...args)
    : () => {};
  billingLogger_ = { log: billingFallbackLog_ }; // eslint-disable-line no-global-assign
}

if (typeof billingSs === 'undefined') {
  function billingSs() {
    if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet) {
      return SpreadsheetApp.getActiveSpreadsheet();
    }
    throw new Error('billingSs is not available');
  }
}

const billingUnpaidHistorySheetName_ = typeof UNPAID_HISTORY_SHEET_NAME === 'string'
  ? UNPAID_HISTORY_SHEET_NAME
  : '未回収履歴';

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

function billingParseTreatmentTimestamp_(rawValue, displayValue) {
  const excelSerialToDate = (value) => {
    const num = Number(value);
    if (!Number.isFinite(num)) return null;
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const millis = excelEpoch.getTime() + Math.round(num * 24 * 60 * 60 * 1000);
    const date = new Date(millis);
    return isNaN(date.getTime()) ? null : date;
  };

  const normalizeDateText = text => {
    if (!text) return '';
    return String(text)
      .replace(/[年\.]/g, '/')
      .replace(/月/g, '/')
      .replace(/日/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  };

  const tryParse = value => {
    if (value instanceof Date && !isNaN(value.getTime())) return value;
    if (typeof value === 'number' && Number.isFinite(value)) {
      const numericDate = excelSerialToDate(value);
      if (numericDate) return numericDate;
    }
    if (value === null || value === undefined) return null;

    const text = String(value).trim();
    if (!text) return null;
    if (/^\d+(\.\d+)?$/.test(text)) {
      const serialDate = excelSerialToDate(text);
      if (serialDate) return serialDate;
    }

    const normalizedText = normalizeDateText(text);
    const parsed = billingParseDateFlexible_(normalizedText);
    return parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : null;
  };

  const parsed = tryParse(rawValue) || tryParse(displayValue) || null;
  if (!parsed) {
    billingLogger_.log('[billing] billingParseTreatmentTimestamp_: failed to parse', rawValue, displayValue);
  }
  return parsed;
}

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

const billingHeaderMapCache_ = {};

function billingGetCachedHeaderMap_(headers) {
  const normalizedHeaders = Array.isArray(headers)
    ? headers.map(header => billingNormalizeHeaderKey_(header)).join('|')
    : '';
  if (!normalizedHeaders) return billingBuildHeaderMap_(headers);
  if (!billingHeaderMapCache_[normalizedHeaders]) {
    billingHeaderMapCache_[normalizedHeaders] = billingBuildHeaderMap_(headers);
  }
  return billingHeaderMapCache_[normalizedHeaders];
}

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

const BILLING_PATIENT_ID_LABELS = typeof PATIENT_ID_LABELS !== 'undefined'
  ? PATIENT_ID_LABELS
  : [
    '施術録番号', '施術録No', '施術録NO',
    '記録番号', 'カルテ番号',
    '患者ID', '患者番号', 'recNo', 'patientId'
  ];

const BILLING_LABELS = typeof LABELS !== 'undefined' ? LABELS : {
  recNo: BILLING_PATIENT_ID_LABELS,
  name: ['名前', '氏名', '氏名（漢字）', '患者名', 'お名前'],
  hospital: ['病院名', '医療機関', '病院'],
  doctor: ['医師', '主治医', '担当医'],
  furigana: ['ﾌﾘｶﾞﾅ', 'ふりがな', 'フリガナ', '氏名（カナ）'],
  birth: ['生年月日', '誕生日', '生年', '生年月'],
  consent: ['同意年月日', '同意日', '同意開始日', '同意開始'],
  consentHandout: ['配布', '配布欄', '配布状況', '配布日', '配布（同意書）'],
  consentContent: ['同意症状', '同意内容', '施術対象疾患', '対象疾患', '対象症状', '同意書内容', '同意記載内容'],
  share: ['負担割合', '負担', '自己負担', '負担率', '負担割', '負担%', '負担％'],
  phone: ['電話', '電話番号', 'TEL', 'Tel']
};

const billingNormalizeEmailKey_ = typeof normalizeEmailKey_ === 'function'
  ? normalizeEmailKey_
  : function normalizeEmailKey_(email) {
    const raw = String(email || '').trim();
    if (!raw) return '';
    const normalized = raw.normalize('NFKC').replace(/\s+/g, '').toLowerCase();
    const plusCollapsed = normalized.replace(/\+[^@]+(?=@)/, '');
    if (plusCollapsed !== normalized) {
      return plusCollapsed;
    }
    if (normalized.includes('@')) {
      return normalized;
    }
    return normalized.replace(/[-_]/g, '');
  };

const billingNormalizeStaffKey_ = typeof normalizeStaffKey_ === 'function'
  ? normalizeStaffKey_
  : function normalizeStaffKeyFallback_(value) {
    const raw = value == null ? '' : String(value).trim();
    if (!raw) return '';

    const normalized = billingNormalizeEmailKey_(raw);
    if (normalized && normalized.includes('@')) {
      return normalized;
    }

    const lowered = raw.normalize('NFKC').toLowerCase();
    const collapsed = lowered
      .replace(/[\s\u3000・･]/g, '')
      .replace(/[ー－ｰ−]/g, '')
      .replace(/[._]/g, '')
      .replace(/[^0-9a-z\u3040-\u30ff\u3400-\u9FFF]/g, '')
      .replace(/-/g, '');
    return collapsed || normalized;
  };

function normalizeStaffKeyCandidates_(candidates) {
  const normalizedKeys = [];
  (candidates || []).forEach(candidate => {
    const normalized = billingNormalizeStaffKey_(candidate);
    if (normalized && normalizedKeys.indexOf(normalized) === -1) {
      normalizedKeys.push(normalized);
    }
  });
  return normalizedKeys;
}

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

const BILLING_PATIENT_RAW_COL_LIMIT = columnLetterToNumber_('BJ');
const BILLING_TREATMENT_SHEET_NAME = '施術録';
const BILLING_PAYMENT_RESULT_SHEET_PREFIX = '入金結果_';
const BILLING_PATIENT_SHEET_NAME = '患者情報';
const BILLING_BANK_SHEET_NAME = '銀行情報';
const BILLING_OVERRIDES_SHEET_NAME = 'BillingOverrides';
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

  billingLogger_.log('[billing] normalizeBillingMonthInput resolved', { input: billingMonth, key, start: start.toISOString(), end: end.toISOString() });

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
  const headerMap = billingGetCachedHeaderMap_(headers);
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
 * - 正式キー: 漢字・カナの双方（存在する方のみ）を NFKC 正規化・空白除去して `kanji::kana` 形式で連結
 * - フォールバック: 正式キーが取れない場合は漢字/カナのいずれか単体を同じ正規化で利用
 * - 最後の手段として patientId を空白除去して利用（全て空の場合は空文字）
 */
function buildBillingNameKey_(record) {
  if (!record) return '';
  const normalizedKanji = normalizeBillingNameKey_(record.nameKanji);
  const normalizedKana = normalizeBillingNameKey_(record.nameKana);
  const fullNameKey = normalizeBillingFullNameKey_(normalizedKanji, normalizedKana);
  if (fullNameKey) return fullNameKey;

  const singleNameKey = normalizeBillingNameKey_(normalizedKanji || normalizedKana);
  if (singleNameKey) return singleNameKey;

  const storedKey = normalizeBillingNameKey_(record._nameKey);
  if (storedKey) return storedKey;

  if (record.patientId != null) {
    const patientIdKey = normalizeBillingNameKey_(record.patientId);
    if (patientIdKey) return patientIdKey;
  }

  return '';
}

function buildPatientIdLookupByNameKey_(patients) {
  return (patients || []).reduce((map, patient) => {
    if (!patient) return map;
    const pid = billingNormalizePatientId_(patient.patientId);
    const nameKey = buildBillingNameKey_(patient);
    if (!pid || !nameKey) return map;

    const entry = map[nameKey] || { ids: [], duplicated: false, patientId: '' };
    if (entry.ids.indexOf(pid) === -1) {
      entry.ids.push(pid);
    }
    entry.duplicated = entry.ids.length > 1;
    entry.patientId = entry.duplicated ? '' : entry.ids[0];
    map[nameKey] = entry;
    return map;
  }, {});
}

function loadBillingStaffDirectory_() {
  const sheet = billingSs().getSheetByName('スタッフ一覧');
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const colName = resolveBillingColumn_(headers, ['名前', '氏名', 'スタッフ名'], '氏名', { fallbackLetter: 'A' });
  const colEmail = resolveBillingColumn_(headers, ['メール', 'メールアドレス', 'email', 'Email'], 'メールアドレス', { fallbackLetter: 'K' });
  const colStaffId = resolveBillingColumn_(headers, ['スタッフID', '担当者ID', 'staffId', 'staffid'], 'スタッフID', {});
  const colNameRoman = resolveBillingColumn_(headers,
    ['ローマ字', 'ﾛｰﾏ字', 'ローマ字氏名', '英字', '英語表記', 'roman', 'romaji', 'romanized'],
    'ローマ字',
    {}
  );
  const colDisabled = resolveBillingColumn_(headers, ['利用停止', '停止', '無効', 'ステータス', 'status'], '利用停止', {});
  if (!colName) return {};

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const displayValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const directory = values.reduce((map, row, idx) => {
    const displayRow = displayValues[idx] || [];
    const name = colName ? String(row[colName - 1] || '').trim() : '';
    if (!name) return map;

    const keys = [];
    const disabledFlag = colDisabled ? normalizeDisabledFlag_(row[colDisabled - 1]) : 0;
    if (disabledFlag === 2) return map;

    if (colEmail) {
      const emailKey = extractNormalizedStaffKey_(row[colEmail - 1], displayRow[colEmail - 1]);
      if (emailKey.normalized) keys.push(emailKey.normalized);
    }
    if (colStaffId) {
      const staffKey = extractNormalizedStaffKey_(row[colStaffId - 1], displayRow[colStaffId - 1]);
      if (staffKey.normalized) keys.push(staffKey.normalized);
    }

    const romanName = colNameRoman ? String(row[colNameRoman - 1] || '').trim() : '';
    const romanDisplay = colNameRoman ? String(displayRow[colNameRoman - 1] || '').trim() : '';
    const nameKeyCandidates = normalizeStaffKeyCandidates_([
      romanName,
      romanDisplay,
      name,
      normalizeBillingNameKey_(name)
    ]);

    nameKeyCandidates.forEach(key => keys.push(key));

    keys.forEach(key => {
      if (key && !map[key]) {
        map[key] = name;
      }
    });
    return map;
  }, {});

  const directoryKeys = Object.keys(directory);
  const staffKeySamples = directoryKeys.slice(0, 20).map(key => ({ key, name: directory[key] }));
  const staffKeyListLog = directoryKeys.slice(0, 200);
  const staffKeyDetailListLog = directoryKeys.slice(0, 200).map(key => ({ key, name: directory[key] }));

  billingLogger_.log('[billing] loadBillingStaffDirectory_: entries=' + directoryKeys.length);
  billingLogger_.log('[billing] loadBillingStaffDirectory_: key samples=' + JSON.stringify(staffKeySamples));
  billingLogger_.log('[billing] loadBillingStaffDirectory_: key list (truncated)=' + JSON.stringify(staffKeyListLog));
  billingLogger_.log('[billing] loadBillingStaffDirectory_: key detail list (truncated)=' + JSON.stringify(staffKeyDetailListLog));
  return directory;
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
  if (String(value).trim() === '自費') return '自費';
  if (typeof value === 'number') {
    if (!isFinite(value)) return 0;
    if (value === 0) return 0;
    if (value < 1) return Math.round(value * 10);
    if (value < 10) return Math.round(value);
    if (value <= 100) return Math.round(value / 10);
  }
  const text = String(value).normalize('NFKC').trim();
  if (!text) return 0;

  const normalized = text.replace(/\s+/g, '').replace('％', '%');
  const hasPercent = normalized.indexOf('%') >= 0;
  const numericText = normalized.replace(/[^0-9.]/g, '');
  const parsed = Number(numericText);
  if (Number.isFinite(parsed)) {
    if (parsed === 0) return 0;
    if (hasPercent) return Math.round(parsed / 10);
    if (parsed < 1) return Math.round(parsed * 10);
    if (parsed < 10) return Math.round(parsed);
    if (parsed <= 100) return Math.round(parsed / 10);
  }

  const ratio = billingNormalizeBurdenRatio_(normalized);
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

function extractEmailFallback_(raw, displayValue) {
  const candidates = [];
  const inputs = [raw, displayValue];
  for (let i = 0; i < inputs.length; i++) {
    const value = inputs[i];
    if (value == null) continue;
    const text = String(value).trim();
    if (!text) continue;
    const matched = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    candidates.push(matched ? matched[0] : text);
  }
  return candidates.find(Boolean) || '';
}

function extractNormalizedStaffKey_(raw, displayValue) {
  const candidate = extractEmailFallback_(raw, displayValue);
  const normalizedList = normalizeStaffKeyCandidates_([candidate]);
  return {
    raw: candidate,
    normalized: normalizedList[0] || ''
  };
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
  const colCreatedByEmail = resolveBillingColumn_(headers,
    ['メール', 'email', 'mail'],
    'メール',
    {}
  );
  const colStaffId = resolveBillingColumn_(headers,
    ['担当者ID', 'スタッフID', 'staffId', 'staffid', 'staff'],
    '担当者ID',
    {});

  const range = sheet.getRange(2, 1, lastRow - 1, width);
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const normalizationDebug = [];
  const invalidDateDebug = [];
  const emptyPidRows = [];
  const staffKeyDebug = [];
  const staffExtractionDebug = [];
  const logs = values.map((row, idx) => {
    const rawPid = row[colPid - 1];
    const pid = billingNormalizePatientId_(rawPid);
    const dateCell = row[colDate - 1];
    const displayRow = displayValues[idx] || [];
    const timestamp = billingParseTreatmentTimestamp_(dateCell, displayRow[colDate - 1]);
    const createdByDisplay = colCreatedBy && displayRow ? displayRow[colCreatedBy - 1] : '';
    const createdByEmailDisplay = colCreatedByEmail && displayRow && displayRow.length >= colCreatedByEmail
      ? displayRow[colCreatedByEmail - 1]
      : '';
    const staffIdRaw = colStaffId ? row[colStaffId - 1] : '';
    const staffIdDisplay = colStaffId && displayRow && displayRow.length >= colStaffId ? displayRow[colStaffId - 1] : '';
    const staffIdInfo = colStaffId ? extractNormalizedStaffKey_(staffIdRaw, staffIdDisplay) : { raw: '', normalized: '' };

    const createdByInfo = colCreatedBy
      ? extractNormalizedStaffKey_(row[colCreatedBy - 1], createdByDisplay)
      : { raw: '', normalized: '' };
    const createdByEmailInfo = colCreatedByEmail
      ? extractNormalizedStaffKey_(row[colCreatedByEmail - 1], createdByEmailDisplay)
      : { raw: '', normalized: '' };

    const selectedInfo = createdByInfo.normalized
      ? createdByInfo
      : (createdByEmailInfo.normalized ? createdByEmailInfo : staffIdInfo);
    const createdByEmail = selectedInfo.raw || createdByInfo.raw || createdByEmailInfo.raw || staffIdInfo.raw;
    const usedNormalized = selectedInfo.normalized || createdByInfo.normalized || createdByEmailInfo.normalized || staffIdInfo.normalized;

    if (String(rawPid || '').trim() && pid && pid !== String(rawPid).trim()) {
      if (normalizationDebug.length < 20) {
        normalizationDebug.push({ rowNumber: idx + 2, rawPid, normalizedPid: pid });
      }
    }
    if (!pid) {
      if (emptyPidRows.length < 20) {
        emptyPidRows.push({ rowNumber: idx + 2, rawPid });
      }
    }
    if (staffKeyDebug.length < 20 && (createdByInfo.normalized || createdByEmailInfo.normalized || staffIdInfo.normalized)) {
      staffKeyDebug.push({
        rowNumber: idx + 2,
        createdByNormalized: createdByInfo.normalized,
        createdByEmailNormalized: createdByEmailInfo.normalized,
        staffIdNormalized: staffIdInfo.normalized,
        usedNormalized,
        usedRaw: createdByEmail
      });
    }
    if (staffExtractionDebug.length < 20 && (createdByInfo.raw || createdByEmailInfo.raw || staffIdInfo.raw)) {
      staffExtractionDebug.push({
        rowNumber: idx + 2,
        createdBy: { raw: createdByInfo.raw, normalized: createdByInfo.normalized },
        createdByEmail: { raw: createdByEmailInfo.raw, normalized: createdByEmailInfo.normalized },
        staffId: { raw: staffIdInfo.raw, normalized: staffIdInfo.normalized },
        used: { raw: createdByEmail, normalized: usedNormalized }
      });
    }
    if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) {
      if (invalidDateDebug.length < 20) {
        invalidDateDebug.push({
          rowNumber: idx + 2,
          patientId: pid,
          rawTimestamp: dateCell,
          displayTimestamp: displayRow[colDate - 1]
        });
      }
    }
    return {
      rowNumber: idx + 2,
      rawPatientId: rawPid,
      patientId: pid,
      timestamp,
      createdByEmail,
      createdByKey: usedNormalized,
      raw: row
    };
  });

  const timestampDebug = logs.map(log => ({
    rowNumber: log.rowNumber,
    patientId: log.patientId,
    timestamp: log.timestamp instanceof Date ? log.timestamp.toISOString() : String(log.timestamp),
    isDate: log.timestamp instanceof Date,
    isValidDate: log.timestamp instanceof Date && !isNaN(log.timestamp.getTime())
  }));
  billingLogger_.log('[billing] loadTreatmentLogs_: timestamps=' + JSON.stringify(timestampDebug));
  billingLogger_.log('[billing] loadTreatmentLogs_: pid normalization samples=' + JSON.stringify(normalizationDebug));
  billingLogger_.log('[billing] loadTreatmentLogs_: invalid date samples=' + JSON.stringify(invalidDateDebug));
  billingLogger_.log('[billing] loadTreatmentLogs_: empty pid rows=' + JSON.stringify(emptyPidRows));
  billingLogger_.log('[billing] loadTreatmentLogs_: staff key samples=' + JSON.stringify(staffKeyDebug));
  billingLogger_.log('[billing] loadTreatmentLogs_: staff extraction samples=' + JSON.stringify(staffExtractionDebug));
  return logs;
}

function buildVisitCountMap_(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const logs = loadTreatmentLogs_();
  const counts = {};
  const staffHistoryByPatient = {};
  let filteredCount = 0;
  const debug = {
    totalLogs: logs.length,
    missingPatientId: 0,
    invalidTimestamp: 0,
    outOfRange: 0,
    counted: 0,
    invalidSamples: [],
    outOfRangeSamples: []
  };
  const skipSamples = { missingPatientId: [], invalidTimestamp: [], outOfRange: [] };
  logs.forEach(log => {
    const pid = log && log.patientId ? billingNormalizePatientId_(log.patientId) : '';
    const ts = log && log.timestamp;
    if (!pid) {
      debug.missingPatientId += 1;
      if (skipSamples.missingPatientId.length < 20) {
        skipSamples.missingPatientId.push({ rowNumber: log.rowNumber, rawPatientId: log.rawPatientId });
      }
      return;
    }
    if (!(ts instanceof Date) || isNaN(ts.getTime())) {
      debug.invalidTimestamp += 1;
      if (debug.invalidSamples.length < 5) {
        debug.invalidSamples.push({ pid, timestamp: ts });
      }
      if (skipSamples.invalidTimestamp.length < 20) {
        skipSamples.invalidTimestamp.push({ rowNumber: log.rowNumber, patientId: pid, timestamp: ts });
      }
      return;
    }
    if (ts < month.start || ts >= month.end) {
      debug.outOfRange += 1;
      if (debug.outOfRangeSamples.length < 5) {
        debug.outOfRangeSamples.push({ pid, timestamp: ts.toISOString() });
      }
      if (skipSamples.outOfRange.length < 20) {
        skipSamples.outOfRange.push({ rowNumber: log.rowNumber, patientId: pid, timestamp: ts.toISOString() });
      }
      return;
    }
    debug.counted += 1;
    filteredCount += 1;
    const current = counts[pid] || { visitCount: 0 };
    current.visitCount += 1;
    counts[pid] = current;

    if (log && log.createdByEmail) {
      const normalizedEmail = log.createdByKey || billingNormalizeStaffKey_(log.createdByEmail);
      if (!normalizedEmail) return;
      if (!staffHistoryByPatient[pid]) {
        staffHistoryByPatient[pid] = {};
      }
      const existing = staffHistoryByPatient[pid][normalizedEmail];
      if (!existing || !existing.timestamp || ts > existing.timestamp) {
        staffHistoryByPatient[pid][normalizedEmail] = { email: log.createdByEmail, timestamp: ts };
      }
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
      .map(entry => entry.email || '')
      .filter(email => !!email);
    map[pid] = sorted;
    return map;
  }, {});
  billingLogger_.log('[billing] buildVisitCountMap_: month range=' + month.start.toISOString() + ' - ' + month.end.toISOString());
  billingLogger_.log('[billing] buildVisitCountMap_: skipped samples=' + JSON.stringify(skipSamples));
  billingLogger_.log('[billing] buildVisitCountMap_: after month filter count=' + filteredCount);
  billingLogger_.log('[billing] buildVisitCountMap_: visitCountMap keys=' + JSON.stringify(Object.keys(counts)));
  billingLogger_.log('[billing] buildVisitCountMap_: staffByPatient size=' + Object.keys(staffByPatient).length);
  billingLogger_.log('[billing] buildVisitCountMap_: staffByPatient mapping (truncated)=' + JSON.stringify(
    Object.keys(staffByPatient)
      .slice(0, 200)
      .map(pid => ({ patientId: pid, staff: staffByPatient[pid] }))
  ));
  billingLogger_.log('[billing] buildVisitCountMap_: debug=' + JSON.stringify(debug));
  return { billingMonth: month.key, counts, staffByPatient, staffHistoryByPatient };
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

function normalizeCarryOverLedgerMonth_(value, fallbackKey) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMM');
  }
  const raw = String(value || '').trim();
  if (!raw && fallbackKey) return fallbackKey;
  const digits = raw.replace(/\D/g, '');
  if (digits.length >= 6) {
    return digits.slice(0, 6);
  }
  return fallbackKey || '';
}

  function ensureCarryOverLedgerSheet_(options) {
    const opts = options || {};
    const meta = opts.meta || {};
    const ss = billingSs();
    let sheet = ss.getSheetByName('CarryOverLedger');
    let didCreate = false;
    const defaultHeader = ['patientId', 'month', 'reason', 'amount', 'createdAt', 'operator'];
    if (!sheet) {
      sheet = ss.insertSheet('CarryOverLedger');
      didCreate = true;
      meta.wasAutoCreated = true;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
    const headerWidth = Math.max(lastCol, defaultHeader.length);
    const headerValues = headerWidth > 0 ? sheet.getRange(1, 1, 1, headerWidth).getDisplayValues()[0] : [];
    const normalizedHeader = (headerValues || []).map(value => String(value || '').trim());
    const hasRequiredHeader = defaultHeader.every(label => normalizedHeader.indexOf(label) >= 0);
    const headerPreview = sheet.getRange(1, 1, 1, 10).getDisplayValues()[0];
    billingLogger_.log('[billing] CarryOverLedger header preview=' + JSON.stringify(headerPreview));

    const shouldInitializeHeader = lastRow === 0 || !hasRequiredHeader;

    if (shouldInitializeHeader) {
      sheet.getRange(1, 1, 1, defaultHeader.length).setValues([defaultHeader]);
      didCreate = true;
      meta.headerInserted = true;
    }

    if (didCreate) {
      billingLogger_.log('[billing] CarryOverLedger auto-created/initialized');
    }
    meta.rowCount = sheet.getLastRow();
    meta.dataRowCount = Math.max(0, meta.rowCount - 1);
    return sheet;
  }

  function loadCarryOverLedgerEntries_(billingMonth, ledgerMeta) {
    const month = normalizeBillingMonthInput(billingMonth);
    const meta = ledgerMeta || {};
    const sheet = ensureCarryOverLedgerSheet_({ meta });
    const lastRow = sheet.getLastRow();
    meta.rowCount = lastRow;
    meta.dataRowCount = Math.max(0, lastRow - 1);
    if (lastRow < 2) return [];
    const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  const colPid = resolveBillingColumn_(headers, ['patientId', '患者ID', '患者番号', 'カルテ番号'], 'patientId', { required: true });
  const colMonth = resolveBillingColumn_(headers, ['month', 'billingMonth', 'ym', '年月', '月'], 'month', {});
  const colReason = resolveBillingColumn_(headers, ['reason', '調整理由', '内訳', 'メモ', 'reasonCode'], 'reason', {});
  const colAmount = resolveBillingColumn_(headers, ['amount', '金額', 'carryOverAmount', '差額', '繰越'], 'amount', { required: true });
  const colCreatedAt = resolveBillingColumn_(headers, ['入力日', '作成日', 'createdAt', '記入日'], '入力日', {});
  const colUser = resolveBillingColumn_(headers, ['操作ユーザー', '作成者', '更新者', 'user', 'operator'], '操作ユーザー', {});

    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const targetMonthNum = Number(month.key);

  return values.map((row, idx) => {
    const pid = billingNormalizePatientId_(row[colPid - 1]);
    if (!pid) return null;
    const amount = normalizeMoneyValue_(row[colAmount - 1]);
    if (!Number.isFinite(amount)) return null;
    const normalizedMonth = normalizeCarryOverLedgerMonth_(colMonth ? row[colMonth - 1] : '', month.key);
    const monthNum = Number(String(normalizedMonth || '').replace(/\D/g, '')) || 0;
    if (targetMonthNum && monthNum && monthNum > targetMonthNum) return null;
    const createdAt = colCreatedAt ? billingParseDateFlexible_(row[colCreatedAt - 1]) : null;
    const reason = colReason ? String(row[colReason - 1] || '').trim() : '';
    const operator = colUser ? String(row[colUser - 1] || '').trim() : '';
    return {
      patientId: pid,
      month: normalizedMonth,
      reason,
      amount,
      createdAt: createdAt || null,
      operator,
      rowNumber: idx + 2
    };
  }).filter(Boolean);
}

function buildSelfPayItemsFromLegacy_(manualSelfPayAmount) {
  const amount = normalizeMoneyValue_(manualSelfPayAmount);
  if (!Number.isFinite(amount) || amount === 0) return [];
  return [{ type: '自費', amount }];
}

function loadBillingOverridesMap_(billingMonth) {
  const month = normalizeBillingMonthInput(billingMonth);
  const sheet = billingSs().getSheetByName(BILLING_OVERRIDES_SHEET_NAME);
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const colYm = resolveBillingColumn_(headers, ['ym', 'billingMonth', 'month'], 'ym', { required: true });
  const colPid = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', 'patientId']), 'patientId', { required: true });
  const colManualUnitPrice = resolveBillingColumn_(headers, ['manualUnitPrice', 'unitPrice', '単価'], 'manualUnitPrice', {});
  const colManualTransport = resolveBillingColumn_(headers, ['manualTransportAmount', 'transportAmount', '交通費'], 'manualTransportAmount', {});
  const colCarryOver = resolveBillingColumn_(headers, ['carryOverAmount', 'carryOver', '未入金額', '繰越'], 'carryOverAmount', {});
  const colAdjustedVisits = resolveBillingColumn_(headers, ['adjustedVisitCount', 'visitCount', '回数'], 'adjustedVisitCount', {});
  const colBurdenRate = resolveBillingColumn_(headers, BILLING_LABELS.share.concat(['burdenRate']), 'burdenRate', {});
  const colManualSelfPay = resolveBillingColumn_(headers, ['manualSelfPayAmount', '自費', 'selfPayAmount'], 'manualSelfPayAmount', {});

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map = {};
  const latestRowByKey = {};
  const cleanupRows = [];
  values.forEach((row, idx) => {
    const ym = String(row[colYm - 1] || '').trim();
    if (!ym || ym !== month.key) return;
    const pid = billingNormalizePatientId_(row[colPid - 1]);
    if (!pid) return;
    const rowNumber = idx + 2;
    const key = ym + '#' + pid;

    const manualUnitPrice = colManualUnitPrice
      ? (row[colManualUnitPrice - 1] === '' || row[colManualUnitPrice - 1] === null
        ? ''
        : normalizeMoneyValue_(row[colManualUnitPrice - 1]))
      : undefined;
    const manualTransportAmount = colManualTransport
      ? (row[colManualTransport - 1] === '' || row[colManualTransport - 1] === null
        ? ''
        : normalizeMoneyValue_(row[colManualTransport - 1]))
      : undefined;
    const carryOverAmount = colCarryOver
      ? (row[colCarryOver - 1] === '' || row[colCarryOver - 1] === null
        ? ''
        : normalizeMoneyValue_(row[colCarryOver - 1]))
      : undefined;
    const adjustedVisitCount = colAdjustedVisits
      ? billingNormalizeVisitCount_(row[colAdjustedVisits - 1])
      : undefined;
    const burdenRate = colBurdenRate
      ? (row[colBurdenRate - 1] === '' || row[colBurdenRate - 1] === null
        ? undefined
        : normalizeBurdenRateInt_(row[colBurdenRate - 1]))
      : undefined;
    const manualSelfPayAmount = colManualSelfPay
      ? (row[colManualSelfPay - 1] === '' || row[colManualSelfPay - 1] === null
        ? ''
        : normalizeMoneyValue_(row[colManualSelfPay - 1]))
      : undefined;

    const hasOverride = [
      manualUnitPrice,
      manualTransportAmount,
      carryOverAmount,
      adjustedVisitCount,
      burdenRate,
      manualSelfPayAmount
    ].some(value => value !== undefined);
    if (!hasOverride) return;

    const override = {
      manualUnitPrice,
      manualTransportAmount,
      carryOverAmount,
      adjustedVisitCount,
      burdenRate,
      manualSelfPayAmount,
      selfPayItems: buildSelfPayItemsFromLegacy_(manualSelfPayAmount)
    };
    const existing = latestRowByKey[key];
    if (existing && existing.rowNumber > rowNumber) {
      cleanupRows.push(rowNumber);
      return;
    }
    if (existing && existing.rowNumber < rowNumber) {
      cleanupRows.push(existing.rowNumber);
    }
    latestRowByKey[key] = { rowNumber, override };
    map[pid] = Object.assign({}, map[pid] || {}, override);
  });

  if (cleanupRows.length && typeof sheet.deleteRow === 'function') {
    cleanupRows
      .sort((a, b) => b - a)
      .forEach(rowNumber => {
        try {
          sheet.deleteRow(rowNumber);
        } catch (err) {
          try {
            billingLogger_.log('[billing] Failed to cleanup duplicate BillingOverrides row=' + rowNumber + ' err=' + err);
          } catch (logErr) {
            // ignore logging errors in non-GAS environments
          }
        }
      });
  }

  return map;
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
  const colUnitPrice = resolveBillingColumn_(
    headers,
    ['単価', '請求単価', '自費単価', '単価(自費)', '単価（自費）', '単価（手動上書き）', '単価(手動上書き)'],
    '単価',
    {}
  );
  const colAddress = resolveBillingColumn_(headers, ['住所', '住所1', '住所２', '住所2', 'address', 'Address'], '住所', {});
  const colPayer = resolveBillingColumn_(headers, ['保険者', '支払区分', '保険/自費', '保険区分種別'], '保険者', {});
  const colBank = resolveBillingColumn_(headers, ['銀行コード', '銀行CD', '銀行番号', 'bankCode'], '銀行コード', { fallbackLetter: 'N' });
  const colBranch = resolveBillingColumn_(headers, ['支店コード', '支店番号', '支店CD', 'branchCode'], '支店コード', { fallbackLetter: 'O' });
  const colAccount = resolveBillingColumn_(headers, ['口座番号', '口座No', '口座NO', 'accountNumber', '口座'], '口座番号', { fallbackLetter: 'Q' });
  const colIsNew = resolveBillingColumn_(headers, ['新規', '新患', 'isNew', '新規フラグ', '新規区分'], '新規区分', { fallbackLetter: 'U' });
  const colCarryOver = resolveBillingColumn_(headers, ['未入金', '未入金額', '未収金', '未収', '繰越', '繰越額', '繰り越し', '差引繰越', '前回未払', '前回未収', 'carryOverAmount'], '未入金額', {});
  const colMedical = resolveBillingColumn_(headers, ['医療助成'], '医療助成', { fallbackLetter: 'AS' });
  const colMedicalSubsidy = colMedical;
  const colTransport = resolveBillingColumn_(
    headers,
    ['交通費', '交通費(手動)', '交通費（手動）', 'transportAmount', 'manualTransportAmount'],
    '交通費',
    {}
  );

  return values.map(row => {
    const pid = billingNormalizePatientId_(row[colPid - 1]);
    if (!pid) return null;
    const insuranceType = colInsurance ? String(row[colInsurance - 1] || '').trim() : '';
    const normalizedBurden = insuranceType === '自費'
      ? '自費'
      : (colBurden ? normalizeBurdenRateInt_(row[colBurden - 1]) : 0);
    const unitPriceRaw = colUnitPrice ? row[colUnitPrice - 1] : '';
    const manualUnitPrice = unitPriceRaw === '' || unitPriceRaw === null
      ? ''
      : normalizeMoneyValue_(unitPriceRaw);
    const transportRaw = colTransport ? row[colTransport - 1] : '';
    const manualTransportAmount = transportRaw === '' || transportRaw === null
      ? ''
      : normalizeMoneyValue_(transportRaw);

    return {
      patientId: pid,
      raw: buildPatientRawObject_(headers, row),
      nameKanji: colName ? String(row[colName - 1] || '').trim() : '',
      nameKana: colKana ? String(row[colKana - 1] || '').trim() : '',
      insuranceType,
      burdenRate: normalizedBurden,
      unitPrice: colUnitPrice ? normalizeMoneyValue_(unitPriceRaw) : 0,
      manualUnitPrice,
      manualTransportAmount,
      address: colAddress ? String(row[colAddress - 1] || '').trim() : '',
      payerType: colPayer ? String(row[colPayer - 1] || '').trim() : '',
      medicalAssistance: colMedical ? normalizeZeroOneFlag_(row[colMedical - 1]) : 0,
      bankCode: colBank ? String(row[colBank - 1] || '').trim() : '',
      branchCode: colBranch ? String(row[colBranch - 1] || '').trim() : '',
      accountNumber: colAccount ? String(row[colAccount - 1] || '').trim() : '',
      isNew: colIsNew ? normalizeZeroOneFlag_(row[colIsNew - 1]) : 0,
      carryOverAmount: colCarryOver ? normalizeMoneyValue_(row[colCarryOver - 1]) : 0,
      medicalSubsidy: colMedicalSubsidy ? normalizeZeroOneFlag_(row[colMedicalSubsidy - 1]) : 0
    };
  }).filter(Boolean);
}

function ensureBankInfoSheet_() {
  const workbook = billingSs();
  let sheet = workbook.getSheetByName(BILLING_BANK_SHEET_NAME);
  if (!sheet) {
    sheet = workbook.insertSheet(BILLING_BANK_SHEET_NAME);
  }

  try {
    const protections = workbook.getProtections(SpreadsheetApp.ProtectionType.SHEET) || [];
    let protection = protections.find(p => {
      const range = p && typeof p.getRange === 'function' ? p.getRange() : null;
      return range && typeof range.getSheet === 'function' && range.getSheet().getName() === sheet.getName();
    });

    if (!protection) {
      protection = sheet.protect();
      protection.setDescription('銀行情報（参照専用マスタ）');
    }

    protection.setWarningOnly(false);
    if (typeof protection.canDomainEdit === 'function' && protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
    const editors = typeof protection.getEditors === 'function' ? protection.getEditors() : [];
    if (editors && editors.length) {
      protection.removeEditors(editors);
    }
  } catch (err) {
    console.warn('[billing] Failed to enforce bank info protection', err);
  }

  return sheet;
}

function getBillingBankRecords(patientRecords) {
  const sheet = ensureBankInfoSheet_();
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
  const colPatientId = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者ID', '患者番号']), '患者ID', {});

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const patientIdLookup = buildPatientIdLookupByNameKey_(patientRecords);
  const records = values.map((row, idx) => {
    const nameKanji = colName ? String(row[colName - 1] || '').trim() : '';
    if (!nameKanji) return null;
    const disabledFlag = colDisabled ? normalizeDisabledFlag_(row[colDisabled - 1]) : 0;
    if (disabledFlag === 2) return null;
    const kanaPrimary = colKana ? String(row[colKana - 1] || '').trim() : '';
    const kanaSecondary = colKanaAlt ? String(row[colKanaAlt - 1] || '').trim() : '';
    const nameKey = buildBillingNameKey_({ nameKanji, nameKana: kanaPrimary || kanaSecondary });
    const rawPatientId = colPatientId ? billingNormalizePatientId_(row[colPatientId - 1]) : '';
    const lookupEntry = nameKey ? patientIdLookup[nameKey] : null;
    const resolvedPatientId = rawPatientId || (lookupEntry && !lookupEntry.duplicated ? lookupEntry.patientId : '');
    const regulationCode = colRegulation ? normalizeMoneyValue_(row[colRegulation - 1]) : 0;
    const isNew = colIsNew ? normalizeZeroOneFlag_(row[colIsNew - 1]) : 0;
    return {
      patientId: resolvedPatientId,
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

  return records;
}

function liftBillingBankProtection_(sheet) {
  const noop = () => {};
  if (!sheet || typeof sheet.getName !== 'function') {
    return { removeProtection: noop, restoreProtection: noop };
  }

  let protection = null;
  try {
    const workbook = billingSs();
    if (workbook && typeof workbook.getProtections === 'function' && typeof SpreadsheetApp !== 'undefined') {
      const protections = workbook.getProtections(SpreadsheetApp.ProtectionType.SHEET) || [];
      protection = protections.find(p => {
        const range = p && typeof p.getRange === 'function' ? p.getRange() : null;
        return range && typeof range.getSheet === 'function' && range.getSheet().getName() === sheet.getName();
      }) || null;
    }
  } catch (err) {
    billingLogger_.log('[billing] liftBillingBankProtection_: failed to locate protection', err);
  }

  const removeProtection = () => {
    if (!protection || typeof protection.remove !== 'function') return;
    try {
      protection.remove();
    } catch (err) {
      billingLogger_.log('[billing] liftBillingBankProtection_: failed to remove protection', err);
    }
  };

  const restoreProtection = () => {
    try {
      ensureBankInfoSheet_();
    } catch (err) {
      billingLogger_.log('[billing] liftBillingBankProtection_: failed to restore protection', err);
    }
  };

  return { removeProtection, restoreProtection };
}

function syncPatientIdToBankInfoSheet_() {
  const sheet = ensureBankInfoSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    billingLogger_.log('[billing] syncPatientIdToBankInfoSheet_: bank sheet is empty');
    return { updated: 0, unresolved: 0 };
  }

  const patients = getBillingPatientRecords();
  const patientIdLookup = buildPatientIdLookupByNameKey_(patients);

  const lastCol = Math.min(sheet.getLastColumn(), sheet.getMaxColumns());
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  const colName = resolveBillingColumn_(headers, BILLING_LABELS.name, '名前', { required: true, fallbackLetter: 'A' });
  const colKana = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', { fallbackLetter: 'B' });
  const colKanaAlt = resolveBillingColumn_(headers, BILLING_LABELS.furigana, 'フリガナ', { fallbackLetter: 'R' });
  const colPatientId = resolveBillingColumn_(headers, BILLING_LABELS.recNo.concat(['患者番号']), '患者ID', { required: true });
  const colDisabled = resolveBillingColumn_(headers, ['利用停止', '停止', '無効', 'ステータス'], '利用停止', { fallbackLetter: 'T' });

  if (!colPatientId) {
    throw new Error('銀行情報シートに患者ID列が見つかりません');
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const patientIdValues = values.map(row => [row[colPatientId - 1]]);
  let updated = 0;
  let unresolved = 0;

  values.forEach((row, idx) => {
    const nameKanji = colName ? String(row[colName - 1] || '').trim() : '';
    const disabledFlag = colDisabled ? normalizeDisabledFlag_(row[colDisabled - 1]) : 0;
    if (disabledFlag === 2) return;
    const kanaPrimary = colKana ? String(row[colKana - 1] || '').trim() : '';
    const kanaSecondary = colKanaAlt ? String(row[colKanaAlt - 1] || '').trim() : '';
    const nameKey = buildBillingNameKey_({ nameKanji, nameKana: kanaPrimary || kanaSecondary });
    const currentPatientId = billingNormalizePatientId_(row[colPatientId - 1]);

    if (!nameKanji && !currentPatientId) {
      unresolved += 1;
      return;
    }

    const patientIdEntry = nameKey ? patientIdLookup[nameKey] : null;
    if (!currentPatientId && patientIdEntry && patientIdEntry.duplicated) {
      unresolved += 1;
      return;
    }

    const resolvedPatientId = currentPatientId || (patientIdEntry ? patientIdEntry.patientId : '');
    if (!resolvedPatientId) {
      unresolved += 1;
      return;
    }
    if (resolvedPatientId !== currentPatientId) {
      patientIdValues[idx][0] = resolvedPatientId;
      updated += 1;
    }
  });

  const { removeProtection, restoreProtection } = liftBillingBankProtection_(sheet);
  try {
    if (updated > 0) {
      removeProtection();
      const targetRange = sheet.getRange(2, colPatientId, lastRow - 1, 1);
      targetRange.setValues(patientIdValues);
    }
  } finally {
    restoreProtection();
  }

  billingLogger_.log('[billing] syncPatientIdToBankInfoSheet_: ' + JSON.stringify({
    updated,
    unresolved
  }));

  return { updated, unresolved };
}

function buildBankLookupByKanji_(bankRecords) {
  return (bankRecords || []).reduce((map, rec) => {
    if (!rec) return map;
    const patientId = billingNormalizePatientId_(rec.patientId);
    if (patientId && !map[patientId]) {
      map[patientId] = rec;
      return map;
    }
    const targetKey = buildBillingNameKey_(rec);
    if (targetKey && !map[targetKey]) {
      map[targetKey] = rec;
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
  const bankRecords = getBillingBankRecords(patientRecords);
  const patientMap = indexByPatientId_(patientRecords);
  const visitCountsResult = buildVisitCountMap_(month);
  const treatmentVisitCounts = visitCountsResult.counts;
  const billingOverrides = loadBillingOverridesMap_(month);
  const billingOverrideFlags = {};
  const mergedPatients = Object.assign({}, patientMap);
  Object.keys(billingOverrides).forEach(pid => {
    const override = billingOverrides[pid];
    if (!override) return;
    const target = Object.assign({}, mergedPatients[pid] || {}, override);
    mergedPatients[pid] = target;
    const flags = billingOverrideFlags[pid] || {};
    if (override.manualUnitPrice !== undefined) flags.unitPrice = true;
    if (override.manualTransportAmount !== undefined) flags.transportAmount = true;
    if (override.carryOverAmount !== undefined) flags.carryOverAmount = true;
    if (override.adjustedVisitCount !== undefined) flags.visitCount = true;
    if (override.burdenRate !== undefined) flags.burdenRate = true;
    if (override.manualSelfPayAmount !== undefined) flags.manualSelfPayAmount = true;
    if (Object.keys(flags).length) {
      billingOverrideFlags[pid] = flags;
    }
    if (override.adjustedVisitCount !== undefined) {
      treatmentVisitCounts[pid] = override.adjustedVisitCount;
    }
  });
  const staffDirectory = loadBillingStaffDirectory_();
  const staffDisplayByPatient = buildStaffDisplayByPatient_(visitCountsResult.staffByPatient || {}, staffDirectory);
    let carryOverLedger = [];
    const carryOverLedgerMeta = {};
    try {
      carryOverLedger = loadCarryOverLedgerEntries_(month, carryOverLedgerMeta);
    } catch (err) {
      billingLogger_.log('[billing] CarryOverLedger missing or inaccessible → using empty fallback: ' + err);
      carryOverLedger = [];
      carryOverLedgerMeta.loadError = String(err);
    }

  const carryOverLedgerByPatient = (carryOverLedger || []).reduce((map, entry) => {
    const pid = billingNormalizePatientId_(entry.patientId);
    if (!pid) return map;
    map[pid] = (map[pid] || 0) + (Number(entry.amount) || 0);
    return map;
  }, {});
  let unpaidHistory = [];
  try {
    unpaidHistory = extractUnpaidBillingHistory(month);
  } catch (err) {
    billingLogger_.log('[billing] unpaidHistory fallback (CarryOverLedger or history missing): ' + err);
    unpaidHistory = [];
  }
  const carryOverFromUnpaid = (unpaidHistory || []).reduce((map, entry) => {
    const pid = billingNormalizePatientId_(entry.patientId);
    if (!pid) return map;
    map[pid] = (map[pid] || 0) + (Number(entry.unpaidAmount) || 0);
    return map;
  }, {});
  const carryOverByPatient = Object.assign({}, carryOverLedgerByPatient);
  Object.keys(carryOverFromUnpaid || {}).forEach(pid => {
    carryOverByPatient[pid] = (carryOverByPatient[pid] || 0) + (carryOverFromUnpaid[pid] || 0);
  });
  const zeroVisitSamples = Object.keys(treatmentVisitCounts || {})
    .filter(pid => {
      const entry = treatmentVisitCounts[pid];
      const visitCount = entry && entry.visitCount != null ? entry.visitCount : entry;
      return !visitCount || Number(visitCount) === 0;
    })
    .slice(0, 20);
  billingLogger_.log('[billing] getBillingSourceData summary=' + JSON.stringify({
    billingMonth: month.key,
    patientCount: patientRecords.length,
    bankRecordCount: bankRecords.length,
    treatmentVisitCountEntries: Object.keys(treatmentVisitCounts || {}).length,
    zeroVisitSamples,
    unpaidHistoryCount: (unpaidHistory || []).length,
      carryOverLedgerCount: (carryOverLedger || []).length,
      carryOverPatients: Object.keys(carryOverByPatient || {}).length,
      staffByPatientCount: Object.keys(visitCountsResult.staffByPatient || {}).length,
      staffDirectorySize: Object.keys(staffDirectory || {}).length
    }));
    billingLogger_.log('[billing] CarryOverLedger status=' + JSON.stringify(carryOverLedgerMeta));
    return {
      billingMonth: month.key,
    month,
    treatmentVisitCounts,
    visitCounts: treatmentVisitCounts,
    patients: mergedPatients,
    patientMap: mergedPatients,
    bankInfoByName: buildBankLookupByKanji_(bankRecords),
    bankStatuses: getBillingPaymentResultsIfExists_(month),
    staffByPatient: visitCountsResult.staffByPatient || {},
    staffHistoryByPatient: visitCountsResult.staffHistoryByPatient || {},
    staffDirectory,
      staffDisplayByPatient,
      unpaidHistory,
      carryOverLedger,
      carryOverLedgerMeta,
      carryOverLedgerByPatient,
      carryOverByPatient,
      billingOverrideFlags
  };
}

function extractUnpaidBillingHistory(targetBillingMonth) {
  const month = normalizeBillingMonthInput(targetBillingMonth);
  const targetKeyNum = Number(month.key);
  const history = extractUnpaidHistoryFromSheet_();
  if (history.length) {
    return history.filter(entry => {
      if (!entry.patientId || !entry.billingMonth) return false;
      if (!entry.entryMonthNum || entry.entryMonthNum >= targetKeyNum) return false;
      return entry.unpaidAmount > 0;
    }).map(entry => {
      const clone = Object.assign({}, entry);
      delete clone.entryMonthNum;
      return clone;
    });
  }

  return extractLegacyUnpaidHistory_(targetKeyNum);
}

function extractUnpaidHistoryFromSheet_() {
  const sheet = billingSs().getSheetByName(billingUnpaidHistorySheetName_);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  return values.map(row => {
    const billingMonth = row[1] ? String(row[1]).trim() : '';
    const entryMonthNum = Number(billingMonth.replace(/\D/g, '')) || 0;
    return {
      patientId: billingNormalizePatientId_(row[0]),
      billingMonth,
      unpaidAmount: Number(row[2]) || 0,
      reason: row[3] || '',
      memo: row[4] || '',
      recordedAt: billingParseDateFlexible_(row[5]) || null,
      entryMonthNum
    };
  }).filter(entry => entry.patientId && entry.billingMonth && entry.unpaidAmount > 0);
}

const BILLING_HISTORY_HEADER_CANDIDATES = typeof BILLING_HISTORY_HEADERS !== 'undefined'
  ? BILLING_HISTORY_HEADERS
  : [
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
    'aggregateUntilMonth'
  ];

function resolveBillingHistoryColumnsFromHeaders_(headers) {
  const map = {};
  (headers || []).forEach((label, idx) => {
    if (BILLING_HISTORY_HEADER_CANDIDATES.indexOf(label) >= 0) {
      map[label] = idx + 1;
    }
  });
  return map;
}

function extractLegacyUnpaidHistory_(targetKeyNum) {
  const sheet = billingSs().getSheetByName('請求履歴');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastColumn = typeof sheet.getLastColumn === 'function' ? sheet.getLastColumn() : 0;
  const colCount = Math.max(lastColumn || 0, 11);
  const headerRange = sheet.getRange(1, 1, 1, colCount);
  const headers = headerRange && typeof headerRange.getDisplayValues === 'function'
    ? headerRange.getDisplayValues()[0]
    : headerRange && typeof headerRange.getValues === 'function'
      ? headerRange.getValues()[0]
      : [];
  const columns = resolveBillingHistoryColumnsFromHeaders_(headers);

  if (!columns.billingMonth) {
    BILLING_HISTORY_HEADER_CANDIDATES.forEach((label, idx) => {
      if (!columns[label] && idx < colCount) {
        columns[label] = idx + 1;
      }
    });
  }
  const values = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();

  return values.map(row => {
    const billingMonth = columns.billingMonth ? String(row[columns.billingMonth - 1] || '').trim() : '';
    const entryMonthNum = Number(billingMonth.replace(/\D/g, '')) || 0;
    return {
      billingMonth,
      patientId: columns.patientId ? billingNormalizePatientId_(row[columns.patientId - 1]) : '',
      nameKanji: columns.nameKanji ? row[columns.nameKanji - 1] || '' : '',
      billingAmount: columns.billingAmount ? Number(row[columns.billingAmount - 1]) || 0 : 0,
      carryOverAmount: columns.carryOverAmount ? Number(row[columns.carryOverAmount - 1]) || 0 : 0,
      grandTotal: columns.grandTotal ? Number(row[columns.grandTotal - 1]) || 0 : 0,
      paidAmount: columns.paidAmount ? Number(row[columns.paidAmount - 1]) || 0 : 0,
      unpaidAmount: columns.unpaidAmount ? Number(row[columns.unpaidAmount - 1]) || 0 : 0,
      bankStatus: columns.bankStatus ? row[columns.bankStatus - 1] || '' : '',
      updatedAt: columns.updatedAt ? billingParseDateFlexible_(row[columns.updatedAt - 1]) || null : null,
      memo: columns.memo ? row[columns.memo - 1] || '' : '',
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
