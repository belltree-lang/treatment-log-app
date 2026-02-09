/***** ── 設定 ─────────────────────────────────*****/
const APP = {
  // Driveに保存するPDFの親フォルダID（空でも可：スプレッドシートと同じ階層に保存）
  PARENT_FOLDER_ID: '1VAv9ZOLB7A__m8ErFDPhFHvhpO21OFPP',
  PAYROLL_PDF_ROOT_FOLDER_ID: '1Jw_QcZ1ph_mi92Y5I2efvpg15VivyV1X',
  // 正本スプレッドシート（患者情報のブック）。空なら「現在のスプレッドシート」を使う
  SSID: '1ajnW9Fuvu0YzUUkfTmw0CrbhrM3lM5tt5OA1dK2_CoQ',
  BASE_FEE_YEN: 4170,
  DOCTOR_REPORT_TEMPLATE_ID: '1mcphwMYaMDVBM0p9MWOv1uMaitNOMPSboi_6F483kZM',
  DOCTOR_REPORT_ROOT_FOLDER_ID: '1CyedMU4jDHsqJqrM234tdhi33W_nn_If',
  // 社内ドメイン制限（空＝無効）
  ALLOWED_DOMAIN: '',   // 例 'belltree1102.com'

  // OpenAI（任意・未設定ならローカル整形へフォールバック）
  OPENAI_ENDPOINT: 'https://api.openai.com/v1/chat/completions',
  OPENAI_MODEL: 'gpt-4o-mini',
};

const SystemPrompt_GenericReport_JP = 'あなたは鍼灸マッサージ院の施術経過を医師・ケアマネ・家族向けに報告する専門アシスタントです。';
const SystemPrompt_DoctorReport_JP = [
  'あなたは鍼灸マッサージ院が医師向けに提出する施術報告書を作成する専門アシスタントです。',
  '以下の法令遵守ルールおよび文調ルールを厳格に守ってください。',
  '',
  '【法令遵守ルール】',
  '・文中で医行為を想起させる語（「治療」「施灸」「刺鍼」「マッサージ」など）は使用禁止です。入力に含まれる場合も必ず「施術」に言い換えてください。',
  '・名詞・動詞いずれも「施術」を用いて記述し、施術の主体は当院スタッフであることを明確にしてください。',
  '',
  '【対象患者に関する前提】',
  '・本報告書の対象は、歩行困難や移動制限を有する在宅療養中の患者様です。',
  '・訪問鍼灸の性質上、日常生活動作（ADL）は何らかの制限があることを前提とします。',
  '',
  '【禁止表現ルール】',
  '・患者様が健常者のように自由な外出・活動を行っていることを示唆する表現は禁止です。',
  ' 例：',
  ' ×「スーパーまで買い物に行けた」',
  ' ×「趣味の登山に出かけた」',
  ' ×「旅行で長距離歩行した」',
  '・これらの表現が入力に含まれていても削除または医療的観察の文脈に修正してください。',
  ' 例：「屋内歩行距離が拡大し、短時間の外出が可能となってきています。」のように、医学的・機能的文脈に置き換えること。',
  '',
  '【文体・記述ルール】',
  '・文章全体は敬体（です・ます調）で統一し、医療文書としての客観性を保ちます。',
  '・主観的・推測的な表現（例：「思います」「感じます」「〜と思われます」など）は使用しないでください。',
  '・施術内容は「当院では〜を実施しております。」の形で現在進行形を基本とします（過去形「いたしました」は避ける）。',
  '・同一文中で「安全に配慮しながら施術を継続してまいります。」を複数回繰り返さないこと。必要に応じて一度のみ使用します。',
  '・段落間は1行空けて構成してください。',
  '・敬称は患者様・医師へ適切に付し、報告書全体を一つの文書として自然な流れにしてください。',
  '・本報告書は「当院（施術提供者）」が「主治医」に対して提出する正式な経過報告書です。',
  '・第三者的・受動的表現（例：「報告が上がっています」「〜と伺っています」「〜とのことです」）は禁止です。',
  '・すべての記述は「当院では」「当院にて」「当院の施術において」を主語として構成してください。',

  '',
  '【構成ルール】',
  '・出力は必ず次の3つの見出しで構成してください（見出し行もそのまま出力します）。',
  '  ■施術の内容・頻度',
  '  ■患者の状態・経過',
  '  ■特記すべき事項',
  '',
  '【内容ルール】',
  '・「施術の内容・頻度」では、冒頭で必ず「頂いている『〇〇』の同意に対して、〜」という形で同意内容を引用し、施術目的と関連付けて記載してください。',
  '・「施術の内容・頻度」では、同意内容 → 施術目的（可動域改善・筋力強化・疼痛緩和など） → 頻度（直近〇か月で〇回の施術を実施、など事実ベース）の順に簡潔に述べてください。',
  '・同意内容が空欄の場合は、「同意内容の記載なし」とせず、施術目的と頻度のみで自然に構成してください。',
  '・「患者の状態・経過」では観察事実を中心に、改善傾向・課題・留意点を簡潔に述べてください。',
  '・医師が判断すべき内容（施術継続の要否、有効性評価、医学的判断）は記載せず、当院が実施している内容と配慮事項のみを事実として記載してください。',
  '・「特記すべき事項」では安全配慮・施術方針・今後の対応などを記載し、最終文は必ず「今後も安全に配慮しながら施術を継続してまいります。」で締めてください（句点を含む）。',
  '',
  '【過去報告書の扱い】',
  '・同一患者様の過去報告書が提示された場合は、内容を参考にしつつ、重複表現を避け、経過の変化を中心にまとめてください。',
  '・前回と同じ内容を繰り返す場合は、文言を自然に言い換えてください。',
  '・報告期間が6か月の場合は、主要な変化点や経過の要約を中心に記述してください。',
  '',
  '【出力形式】',
  '・出力は3つの見出しを含む本文のみで行い、挨拶文や署名は不要です。',
  '・日本語のみで出力してください。'
].join('\n');
const AI_REPORT_SHEET_HEADER = ['TS','患者ID','範囲','対象','対象キー','本文','status','special','期間（月）','参照元レポートID','生成方式'];

const AUX_SHEETS_INIT_KEY = 'AUX_SHEETS_INIT_V202503';
const CONSENT_NEWS_META_STANDARDIZED_KEY = 'CONSENT_NEWS_META_STANDARDIZED_V1';
const PATIENT_CACHE_TTL_SECONDS = 90;
const PATIENT_CACHE_KEYS = {
  header: pid => 'patient:header:' + normId_(pid),
  news: pid => 'patient:news:' + normId_(pid),
  treatments: pid => 'patient:treatments:' + normId_(pid),
  reports: pid => 'patient:reports:' + normId_(pid),
  latestTreatmentRow: pid => 'patient:latestTreatRow:' + normId_(pid),
};
const GLOBAL_NEWS_CACHE_KEY = 'patient:news:__global__';
const DOCTOR_REPORT_HANDOVER_WINDOW_DAYS = 30;

const TREATMENT_SHEET_HEADER = [
  'タイムスタンプ',
  '施術録番号',
  '所見',
  'メール',
  '最終確認',
  '名前',
  'treatmentId',
  '施術時間区分',
  '換算人数',
  '新規対応人数',
  '総換算人数',
  '勤怠反映フラグ',
  'NewsConsentDismissed'
];

const TREATMENT_CATEGORY_DEFINITIONS = {
  insurance30: { label: '30分施術（保険）', allowEmptyPatientId: false },
  self30:      { label: '30分施術（自費）', allowEmptyPatientId: false },
  self60:      { label: '60分施術（完全自費）', allowEmptyPatientId: false },
  mixed:       { label: '60分施術（保険＋自費）', allowEmptyPatientId: false },
  new:         { label: '新規', allowEmptyPatientId: true }
};

const TREATMENT_CATEGORY_ATTENDANCE_METRICS = {
  insurance30: { convertedCount: 1, newPatientCount: 0 },
  self30:      { convertedCount: 1, newPatientCount: 0 },
  self60:      { convertedCount: 2, newPatientCount: 0 },
  mixed:       { convertedCount: 1.5, newPatientCount: 0 },
  new:         { convertedCount: 1, newPatientCount: 1 }
};

const TREATMENT_CATEGORY_LABEL_TO_KEY = Object.keys(TREATMENT_CATEGORY_DEFINITIONS).reduce((map, key) => {
  const def = TREATMENT_CATEGORY_DEFINITIONS[key];
  if (def && def.label) {
    map[def.label] = key;
  }
  return map;
}, {});

const TREATMENT_CATEGORY_ATTENDANCE_GROUP = {
  insurance30: 'insurance',
  self30: 'self',
  self60: 'self',
  mixed: 'mixed',
  new: 'new'
};

const VISIT_ATTENDANCE_SHEET_NAME = 'VisitAttendance';
const VISIT_ATTENDANCE_SHEET_HEADER = ['日付','メール','出勤','退勤','勤務時間','休憩','種別内訳','自動反映フラグ','leaveType','isHourlyStaff','isDailyStaff','source'];
const VISIT_ATTENDANCE_AUTO_FLAG_VALUE = 'auto';
const VISIT_ATTENDANCE_WORK_START_MINUTES = 9 * 60;
const VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES = 18 * 60;
const VISIT_ATTENDANCE_ROUNDING_MINUTES = 15;
const VISIT_ATTENDANCE_REQUEST_SHEET_NAME = 'VisitAttendanceRequests';
const VISIT_ATTENDANCE_REQUEST_SHEET_HEADER = [
  'ID',
  'TS',
  '申請者',
  '対象メール',
  '対象日',
  '出勤',
  '退勤',
  '休憩(分)',
  '申請メモ',
  '状態',
  '状態更新',
  '対応者',
  '対応メモ',
  '原データ',
  '申請種別'
];
const VISIT_ATTENDANCE_REQUEST_TYPE_CORRECTION = 'correction';
const VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE = 'paidLeave';
const VISIT_ATTENDANCE_STAFF_SHEET_NAME = 'VisitAttendanceStaff';
const VISIT_ATTENDANCE_STAFF_SHEET_HEADER = ['メール','表示名','年間有給付与日数','雇用区分','AlbyteスタッフID','標準勤務時間(分)','給与従業員ID','拠点'];
const DEFAULT_ANNUAL_PAID_LEAVE_DAYS = 10;
const PAID_LEAVE_DEFAULT_WORK_MINUTES = 8 * 60;
const VISIT_ATTENDANCE_DEFAULT_SHIFT_MINUTES = 8 * 60;
const PAID_LEAVE_HALF_MINIMUM_MINUTES = 2 * 60;
const PAID_LEAVE_EMPLOYMENT_LABELS = Object.freeze({
  employee: '社員',
  parttime: 'アルバイト',
  daily: '日給'
});
const PAID_LEAVE_TYPE_LABELS = Object.freeze({
  full: '全日',
  amHalf: '半日（午前）',
  pmHalf: '半日（午後）'
});

const ALBYTE_STAFF_SHEET_NAME = 'AlbyteStaff';
const ALBYTE_ATTENDANCE_SHEET_NAME = 'AlbyteAttendance';
const ALBYTE_STAFF_SHEET_HEADER = ['ID','名前','PIN','ロック中','連続失敗','最終ログイン','更新TS','スタッフ種別','基準退勤'];
const ALBYTE_ATTENDANCE_SHEET_HEADER = [
  'ID',
  'スタッフID',
  'スタッフ名',
  '日付',
  '出勤',
  '退勤',
  '休憩(分)',
  '備考',
  '自動補正',
  '打刻ログ',
  '作成TS',
  '更新TS'
];
const ALBYTE_SHIFT_SHEET_NAME = 'AlbyteShifts';
const ALBYTE_SHIFT_SHEET_HEADER = ['ID','日付','スタッフID','スタッフ名','開始','終了','メモ','更新TS'];
const ALBYTE_MAX_PIN_ATTEMPTS = 5;
const ALBYTE_SESSION_SECRET_PROPERTY_KEY = 'ALBYTE_SESSION_SECRET';
const ALBYTE_SESSION_TTL_MILLIS = 1000 * 60 * 60 * 12;
const ALBYTE_BREAK_MINUTES_PRESETS = Object.freeze([30, 45, 60, 90, 120, 180]);
const ALBYTE_BREAK_STEP_MINUTES = 15;
const ALBYTE_MAX_BREAK_MINUTES = 180;
const ALBYTE_DAILY_OVERTIME_ROUNDING_MINUTES = 15;
const ALBYTE_HOURLY_WAGE_PROPERTY_KEYS = Object.freeze(['ALBYTE_HOURLY_WAGE', 'albyteHourlyWage']);

const PAYROLL_EMPLOYEE_SHEET_NAME = 'PayrollEmployees';
const PAYROLL_EMPLOYEE_SHEET_HEADER = ['従業員ID','氏名','メール','拠点','雇用区分','基本給','時給','個別加算','役職/等級','資格手当','車両手当','社宅控除','住民税','源泉徴収','交通費区分','交通費額','歩合ロジック','メモ','更新日時','扶養人数','甲乙区分','雇用期間区分'];
const PAYROLL_EMPLOYEE_COLUMNS = Object.freeze({
  id: 0,
  name: 1,
  email: 2,
  base: 3,
  employmentType: 4,
  baseSalary: 5,
  hourlyWage: 6,
  personalAllowance: 7,
  grade: 8,
  qualificationAllowance: 9,
  vehicleAllowance: 10,
  housingDeduction: 11,
  municipalTax: 12,
  withholding: 13,
  transportationType: 14,
  transportationAmount: 15,
  commissionLogic: 16,
  note: 17,
  updatedAt: 18,
  dependentCount: 19,
  withholdingCategory: 20,
  withholdingPeriodType: 21
});
const PAYROLL_EMPLOYEE_COLUMN_INDEX = Object.freeze(Object.keys(PAYROLL_EMPLOYEE_COLUMNS).reduce((map, key) => {
  map[key] = PAYROLL_EMPLOYEE_COLUMNS[key] + 1;
  return map;
}, {}));
const PAYROLL_EMPLOYMENT_LABELS = Object.freeze({
  employee: '正社員',
  parttime: 'アルバイト',
  contractor: '業務委託'
});
const PAYROLL_TRANSPORTATION_LABELS = Object.freeze({
  fixed: '固定',
  actual: '実費',
  none: 'なし'
});
const PAYROLL_COMMISSION_LABELS = Object.freeze({
  legacy: '既存',
  horiguchi: '堀口以降'
});
const PAYROLL_COMMISSION_RULES = Object.freeze({
  legacy: { monthlyThreshold: 7, amount: 1250 },
  horiguchi: { weeklyThreshold: 30, amount: 1250 }
});
const PAYROLL_WITHHOLDING_LABELS = Object.freeze({
  required: 'あり',
  none: 'なし（個人事業主扱い）'
});
const PAYROLL_WITHHOLDING_TAX_RATE = 0.1021;
const PAYROLL_INCOME_TAX_SHEET_NAME = '所得税税額表';
const PAYROLL_INCOME_TAX_CACHE_KEY = 'PAYROLL_INCOME_TAX_SHEET_CACHE';
const PAYROLL_INCOME_TAX_CACHE_TTL_SECONDS = 60 * 60; // 1時間
const PAYROLL_WITHHOLDING_CATEGORY_LABELS = Object.freeze({
  ko: '甲欄',
  otsu: '乙欄'
});
const PAYROLL_WITHHOLDING_PERIOD_LABELS = Object.freeze({
  monthly: '通常（月額）',
  daily: '日額扱い'
});
const PAYROLL_INCOME_TAX_TABLE_URL_PROPERTY_KEY = 'PAYROLL_INCOME_TAX_TABLE_URL';
const PAYROLL_INCOME_TAX_TABLE_CACHE_PROPERTY_KEY = 'PAYROLL_INCOME_TAX_TABLE_CACHE';
const WITHHOLDING_TAX_TABLE_KO_PROPERTY_KEY = 'WITHHOLDING_TAX_TABLE_KO';
const WITHHOLDING_TAX_TABLE_OTSU_PROPERTY_KEY = 'WITHHOLDING_TAX_TABLE_OTSU';
const WITHHOLDING_TAX_TABLE_META_PROPERTY_KEY = 'WITHHOLDING_TAX_TABLE_META';
const PAYROLL_GRADE_SHEET_NAME = 'PayrollGrades';
const PAYROLL_GRADE_SHEET_HEADER = ['グレードID','役職/等級','手当額','メモ','更新日時'];
const PAYROLL_GRADE_COLUMNS = Object.freeze({
  id: 0,
  name: 1,
  amount: 2,
  note: 3,
  updatedAt: 4
});
const PAYROLL_GRADE_COLUMN_INDEX = Object.freeze(Object.keys(PAYROLL_GRADE_COLUMNS).reduce((map, key) => {
  map[key] = PAYROLL_GRADE_COLUMNS[key] + 1;
  return map;
}, {}));
const PAYROLL_GRADE_DEFAULTS = Object.freeze([
  { name: '施設長' },
  { name: '副施設長' },
  { name: '院長' },
  { name: '副院長' }
]);
const PAYROLL_SOCIAL_INSURANCE_STANDARD_SHEET_NAME = 'PayrollInsuranceStandards';
const PAYROLL_SOCIAL_INSURANCE_STANDARD_HEADER = ['標準報酬ID','等級','標準報酬月額','報酬下限','報酬上限','メモ','更新日時'];
const PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS = Object.freeze({
  id: 0,
  grade: 1,
  monthlyAmount: 2,
  lowerBound: 3,
  upperBound: 4,
  note: 5,
  updatedAt: 6
});
const PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMN_INDEX = Object.freeze(Object.keys(PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS).reduce((map, key) => {
  map[key] = PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS[key] + 1;
  return map;
}, {}));
const PAYROLL_SOCIAL_INSURANCE_OVERRIDE_SHEET_NAME = 'PayrollInsuranceOverrides';
const PAYROLL_SOCIAL_INSURANCE_OVERRIDE_HEADER = ['上書きID','従業員ID','年月','等級','標準報酬月額','メモ','更新日時'];
const PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS = Object.freeze({
  id: 0,
  employeeId: 1,
  monthKey: 2,
  grade: 3,
  monthlyAmount: 4,
  note: 5,
  updatedAt: 6
});
const PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMN_INDEX = Object.freeze(Object.keys(PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS).reduce((map, key) => {
  map[key] = PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS[key] + 1;
  return map;
}, {}));
const PAYROLL_ROLE_SHEET_NAME = 'PayrollRoles';
const PAYROLL_ROLE_SHEET_HEADER = ['メール','ロール','拠点'];
const PAYROLL_OWNER_EMAILS = Object.freeze([
  'belltree@belltree1102.com',
  'suzuki@belltree1102.com'
]);
const PAYROLL_SOCIAL_INSURANCE_RATE_PROPERTY_KEY = 'payroll_social_insurance_rates';
const PAYROLL_PDF_ROOT_FOLDER_PROPERTY_KEY = 'PAYROLL_PDF_ROOT_FOLDER_ID';
const PAYROLL_SOCIAL_INSURANCE_RATE_DEFAULTS = Object.freeze({
  healthEmployee: 0.0495,
  healthEmployer: 0.0495,
  pensionEmployee: 0.0915,
  pensionEmployer: 0.0915,
  nursingEmployee: 0.0045,
  nursingEmployer: 0.0045,
  childEmployee: 0.0018,
  childEmployer: 0.0018,
  employmentEmployee: 0,
  employmentEmployer: 0
});

const PAYROLL_PAYOUT_TYPE_LABELS = Object.freeze({
  salary: '通常給与',
  bonus: '賞与',
  yearEndAdjustment: '年末調整',
  withholdingCertificate: '源泉徴収票'
});

const PAYROLL_PAYOUT_EVENT_SHEET_NAME = 'PayrollPayoutEvents';
const PAYROLL_PAYOUT_EVENT_HEADER = [
  'イベントID',
  '従業員ID',
  '支給種別',
  '対象年度',
  '対象月',
  '期間開始',
  '期間終了',
  '支給日',
  'タイトル',
  '状態',
  '明細JSON',
  '社会保険JSON',
  '調整JSON',
  'メタJSON',
  '更新日時'
];
const PAYROLL_PAYOUT_EVENT_COLUMNS = Object.freeze({
  id: 0,
  employeeId: 1,
  payoutType: 2,
  fiscalYear: 3,
  monthKey: 4,
  periodStart: 5,
  periodEnd: 6,
  payDate: 7,
  title: 8,
  status: 9,
  detailsJson: 10,
  insuranceJson: 11,
  adjustmentJson: 12,
  metadataJson: 13,
  updatedAt: 14
});
const PAYROLL_PAYOUT_EVENT_COLUMN_INDEX = Object.freeze(Object.keys(PAYROLL_PAYOUT_EVENT_COLUMNS).reduce((map, key) => {
  map[key] = PAYROLL_PAYOUT_EVENT_COLUMNS[key] + 1;
  return map;
}, {}));

const PAYROLL_ANNUAL_SUMMARY_SHEET_NAME = 'PayrollAnnualSummaries';
const PAYROLL_ANNUAL_SUMMARY_HEADER = [
  '集計ID',
  '従業員ID',
  '対象年度',
  '課税支給額',
  '非課税支給額',
  '社会保険料',
  '雇用保険料',
  '源泉所得税',
  '住民税',
  '年末調整額',
  '賞与総額',
  '支給回数',
  'summaryJSON',
  'メタJSON',
  '更新日時'
];
const PAYROLL_ANNUAL_SUMMARY_COLUMNS = Object.freeze({
  id: 0,
  employeeId: 1,
  fiscalYear: 2,
  taxableAmount: 3,
  nonTaxableAmount: 4,
  socialInsurance: 5,
  employmentInsurance: 6,
  withholdingTax: 7,
  municipalTax: 8,
  yearEndAdjustment: 9,
  bonusAmount: 10,
  payoutCount: 11,
  summaryJson: 12,
  metadataJson: 13,
  updatedAt: 14
});
const PAYROLL_ANNUAL_SUMMARY_COLUMN_INDEX = Object.freeze(Object.keys(PAYROLL_ANNUAL_SUMMARY_COLUMNS).reduce((map, key) => {
  map[key] = PAYROLL_ANNUAL_SUMMARY_COLUMNS[key] + 1;
  return map;
}, {}));

const ALBYTE_STAFF_COLUMNS = Object.freeze({
  id: 0,
  name: 1,
  pin: 2,
  locked: 3,
  failCount: 4,
  lastLogin: 5,
  updatedAt: 6,
  staffType: 7,
  shiftEndTime: 8
});

const ALBYTE_STAFF_COLUMN_INDEX = Object.freeze(Object.keys(ALBYTE_STAFF_COLUMNS).reduce((map, key) => {
  map[key] = ALBYTE_STAFF_COLUMNS[key] + 1;
  return map;
}, {}));

const ALBYTE_ATTENDANCE_COLUMNS = Object.freeze({
  id: 0,
  staffId: 1,
  staffName: 2,
  date: 3,
  clockIn: 4,
  clockOut: 5,
  breakMinutes: 6,
  note: 7,
  autoFlag: 8,
  log: 9,
  createdAt: 10,
  updatedAt: 11
});

const ALBYTE_ATTENDANCE_COLUMN_INDEX = Object.freeze(Object.keys(ALBYTE_ATTENDANCE_COLUMNS).reduce((map, key) => {
  map[key] = ALBYTE_ATTENDANCE_COLUMNS[key] + 1;
  return map;
}, {}));

const ALBYTE_SHIFT_COLUMNS = Object.freeze({
  id: 0,
  date: 1,
  staffId: 2,
  staffName: 3,
  start: 4,
  end: 5,
  note: 6,
  updatedAt: 7
});

const ALBYTE_SHIFT_COLUMN_INDEX = Object.freeze(Object.keys(ALBYTE_SHIFT_COLUMNS).reduce((map, key) => {
  map[key] = ALBYTE_SHIFT_COLUMNS[key] + 1;
  return map;
}, {}));

function getAlbyteHourlyWage_(){
  for (let i = 0; i < ALBYTE_HOURLY_WAGE_PROPERTY_KEYS.length; i++) {
    const key = ALBYTE_HOURLY_WAGE_PROPERTY_KEYS[i];
    const raw = getConfig(key);
    if (raw == null || raw === '') continue;
    const value = Number(raw);
    if (Number.isFinite(value) && value > 0) {
      return value;
    }
  }
  return null;
}

function getScriptCache_(){
  try {
    return CacheService.getScriptCache();
  } catch (e) {
    Logger.log('[cache] CacheService unavailable: ' + (e && e.message ? e.message : e));
    return null;
  }
}

function cacheFetch_(key, fetchFn, ttlSeconds){
  const cache = getScriptCache_();
  if (!cache || !key || typeof fetchFn !== 'function') {
    return fetchFn ? fetchFn() : null;
  }

  try {
    const hit = cache.get(key);
    if (hit != null && hit !== '') {
      return JSON.parse(hit);
    }
  } catch (err) {
    Logger.log('[cache] read miss (' + key + '): ' + (err && err.message ? err.message : err));
  }

  const fresh = fetchFn();
  if (fresh === undefined) return fresh;

  try {
    cache.put(key, JSON.stringify(fresh), Math.max(5, ttlSeconds || PATIENT_CACHE_TTL_SECONDS));
  } catch (err) {
    Logger.log('[cache] write fail (' + key + '): ' + (err && err.message ? err.message : err));
  }
  return fresh;
}

function invalidateCacheKeys_(keys){
  if (!Array.isArray(keys) || !keys.length) return;
  const filtered = keys.filter(Boolean);
  if (!filtered.length) return;
  const cache = getScriptCache_();
  if (!cache) return;
  try {
    cache.removeAll(filtered);
  } catch (err) {
    Logger.log('[cache] remove fail: ' + (err && err.message ? err.message : err));
  }
}

function invalidatePatientCaches_(pidOrList, scope){
  if (Array.isArray(pidOrList)) {
    const allKeys = [];
    pidOrList.forEach(id => {
      const keys = collectPatientCacheKeys_(id, scope);
      if (keys.length) allKeys.push.apply(allKeys, keys);
    });
    invalidateCacheKeys_(allKeys);
    return;
  }
  const keys = collectPatientCacheKeys_(pidOrList, scope);
  invalidateCacheKeys_(keys);
}

function collectPatientCacheKeys_(pid, scope){
  const normalized = normId_(pid);
  if (!normalized) return [];
  const applyAll = !scope;
  const keys = [];
  if (applyAll || scope.header) keys.push(PATIENT_CACHE_KEYS.header(normalized));
  if (applyAll || scope.news) keys.push(PATIENT_CACHE_KEYS.news(normalized));
  if (applyAll || scope.treatments) keys.push(PATIENT_CACHE_KEYS.treatments(normalized));
  if (applyAll || scope.reports) keys.push(PATIENT_CACHE_KEYS.reports(normalized));
  if (applyAll || scope.latestTreatmentRow) keys.push(PATIENT_CACHE_KEYS.latestTreatmentRow(normalized));
  return keys;
}

function invalidateGlobalNewsCache_(){
  invalidateCacheKeys_([GLOBAL_NEWS_CACHE_KEY]);
}

function toBooleanFromCell_(cell){
  if (cell == null) return false;
  if (typeof cell === 'boolean') return cell;
  if (typeof cell === 'number') return cell !== 0;
  if (cell instanceof Date) return true;
  const text = String(cell || '').trim().toLowerCase();
  if (!text) return false;
  if (text === 'false' || text === '0' || text === 'no' || text === 'off') return false;
  return true;
}

/***** 先頭行（見出し）の揺れに耐えるためのラベル候補群 *****/
const PATIENT_ID_LABELS = [
  '施術録番号', '施術録No', '施術録NO', '記録番号', 'カルテ番号', '患者ID', '患者番号', 'recNo', 'patientId'
];
const LABELS = {
  recNo:     PATIENT_ID_LABELS,
  name:      ['名前','氏名','患者名','お名前'],
  hospital:  ['病院名','医療機関','病院'],
  doctor:    ['医師','主治医','担当医'],
  furigana:  ['ﾌﾘｶﾞﾅ','ふりがな','フリガナ'],
  birth:     ['生年月日','誕生日','生年','生年月'],
  consent:   ['同意年月日','同意日','同意開始日','同意開始'],
  consentHandout: ['配布','配布欄','配布状況','配布日','配布（同意書）'],
  consentContent: ['同意症状','同意内容','施術対象疾患','対象疾患','対象症状','同意書内容','同意記載内容'],
  share:     ['負担割合','負担','自己負担','負担率','負担割','負担%','負担％'],
  phone:     ['電話','電話番号','TEL','Tel']
};

// 固定列のフォールバック（どうしても見出しが見つからない時はこれを使う）
const PATIENT_COLS_FIXED = {
  recNo:    3,   // 施術録番号
  name:     4,   // 名前
  hospital: 5,   // 病院名
  furigana: 6,   // ﾌﾘｶﾞﾅ
  birth:    7,   // 生年月日
  doctor:  26,   // 医師
  consent: 28,   // 同意年月日
  consentHandout: 54, // 配布（同意書取得日）
  consentContent: 25, // 同意症状（Y列）
  phone:   32,   // 電話
  share:   47    // 負担割合
};

/***** スプレッドシート参照ユーティリティ *****/

/***** 権限制限（社内ドメインのみ） *****/
function assertDomain_() {
  if (!APP.ALLOWED_DOMAIN) return;
  const email = (Session.getActiveUser() || {}).getEmail() || '';
  if (!email.endsWith('@' + APP.ALLOWED_DOMAIN)) {
    throw new Error('権限がありません（社内ドメインのみ）');
  }
}

const AFTER_TREATMENT_TRIGGER_KEY = 'AFTER_JOBS_TRIGGER_TS';

function scheduleAfterTreatmentJobTrigger_(options){
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const lastScheduled = Number(props.getProperty(AFTER_TREATMENT_TRIGGER_KEY) || '0');
  const minInterval = options && typeof options.minIntervalMs === 'number' ? options.minIntervalMs : 5000;
  if (!options || !options.force) {
    if (lastScheduled && now - lastScheduled < minInterval) {
      return;
    }
  }

  const delayMs = options && options.delayMs != null ? options.delayMs : 5000;
  const delaySeconds = Math.max(1, Math.round(delayMs / 1000));
  try {
    ScriptApp.newTrigger('afterTreatmentJob')
      .timeBased()
      .after(delaySeconds * 1000)
      .create();
    props.setProperty(AFTER_TREATMENT_TRIGGER_KEY, String(now));
  } catch (err) {
    const message = err && err.message ? err.message : err;
    Logger.log('[queueAfterTreatmentJob] Failed to schedule trigger: ' + message);
    if (!options || !options.skipFallback) {
      try {
        ScriptApp.newTrigger('afterTreatmentJob')
          .timeBased()
          .after(60 * 1000)
          .create();
      } catch (fallbackErr) {
        const fallbackMessage = fallbackErr && fallbackErr.message ? fallbackErr.message : fallbackErr;
        Logger.log('[queueAfterTreatmentJob] Fallback trigger failed: ' + fallbackMessage);
      }
    }
  }
}
/***** 補助タブの用意（不足時に自動生成＋ヘッダ挿入） *****/
function ensureAuxSheets_(options) {
  const props = PropertiesService.getScriptProperties();
  const force = options && options.force;
  if (!force && props.getProperty(AUX_SHEETS_INIT_KEY) === '1') {
    return;
  }

  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    locked = lock.tryLock(5000);
  } catch (e) {
    locked = false;
  }

  try {
    if (!force && props.getProperty(AUX_SHEETS_INIT_KEY) === '1') {
      return;
    }

    const wb = ss();
    const need = ['施術録','患者情報','News','フラグ','予定','操作ログ','定型文','添付索引','年次確認','ダッシュボード','AI報告書', VISIT_ATTENDANCE_SHEET_NAME, ALBYTE_ATTENDANCE_SHEET_NAME, ALBYTE_STAFF_SHEET_NAME, PAYROLL_EMPLOYEE_SHEET_NAME, PAYROLL_GRADE_SHEET_NAME, PAYROLL_SOCIAL_INSURANCE_STANDARD_SHEET_NAME, PAYROLL_SOCIAL_INSURANCE_OVERRIDE_SHEET_NAME, PAYROLL_ROLE_SHEET_NAME, PAYROLL_PAYOUT_EVENT_SHEET_NAME, PAYROLL_ANNUAL_SUMMARY_SHEET_NAME];
    need.forEach(n => { if (!wb.getSheetByName(n)) wb.insertSheet(n); });

    const ensureHeader = (name, header) => {
      const s = wb.getSheetByName(name);
      if (s.getLastRow() === 0) s.appendRow(header);
    };

    // 既存タブ
    ensureHeader('施術録',   TREATMENT_SHEET_HEADER);
    ensureHeader('News',     ['TS','患者ID','種別','メッセージ','cleared','meta','dismissed']);

    const upgradeHeader = (sheetName, header) => {
      const sheet = wb.getSheetByName(sheetName);
      if (!sheet) return;
      const needed = header.length;
      if (sheet.getMaxColumns() < needed) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
      }
      const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
      const mismatch = current.length < needed || header.some((label, idx) => String(current[idx] || '') !== label);
      if (mismatch) {
        sheet.getRange(1, 1, 1, needed).setValues([header]);
      }
    };

    upgradeHeader('施術録', TREATMENT_SHEET_HEADER);
    upgradeHeader('News',   ['TS','患者ID','種別','メッセージ','cleared','meta','dismissed']);
    upgradeHeader('AI報告書', AI_REPORT_SHEET_HEADER);
    upgradeHeader(VISIT_ATTENDANCE_SHEET_NAME, VISIT_ATTENDANCE_SHEET_HEADER);
    upgradeHeader(ALBYTE_ATTENDANCE_SHEET_NAME, ALBYTE_ATTENDANCE_SHEET_HEADER);
    upgradeHeader(ALBYTE_STAFF_SHEET_NAME, ALBYTE_STAFF_SHEET_HEADER);
    upgradeHeader(PAYROLL_EMPLOYEE_SHEET_NAME, PAYROLL_EMPLOYEE_SHEET_HEADER);
    upgradeHeader(PAYROLL_GRADE_SHEET_NAME, PAYROLL_GRADE_SHEET_HEADER);
    upgradeHeader(PAYROLL_SOCIAL_INSURANCE_STANDARD_SHEET_NAME, PAYROLL_SOCIAL_INSURANCE_STANDARD_HEADER);
    upgradeHeader(PAYROLL_SOCIAL_INSURANCE_OVERRIDE_SHEET_NAME, PAYROLL_SOCIAL_INSURANCE_OVERRIDE_HEADER);
    upgradeHeader(PAYROLL_ROLE_SHEET_NAME, PAYROLL_ROLE_SHEET_HEADER);
    upgradeHeader(PAYROLL_PAYOUT_EVENT_SHEET_NAME, PAYROLL_PAYOUT_EVENT_HEADER);
    upgradeHeader(PAYROLL_ANNUAL_SUMMARY_SHEET_NAME, PAYROLL_ANNUAL_SUMMARY_HEADER);
    ensureHeader('フラグ',   ['患者ID','status','pauseUntil']);
    ensureHeader('予定',     ['患者ID','種別','予定日','登録者']);
    ensureHeader('操作ログ', ['TS','操作','患者ID','詳細','実行者']);
    ensureHeader('定型文',   ['カテゴリ','ラベル','文章']);
    ensureHeader('添付索引', ['TS','患者ID','月','ファイル名','FileId','種別','登録者']);
    ensureHeader('AI報告書', AI_REPORT_SHEET_HEADER);

    // 年次確認タブ（未作成時はヘッダだけ用意）
    ensureHeader('年次確認', ['患者ID','年','確認日','担当者メール']);

    // ダッシュボード（Index）タブ
    ensureHeader('ダッシュボード', [
      '患者ID','氏名','同意年月日','次回期限','期限ステータス',
      '担当者(60d)','最終施術日','年次要確認','休止','ミュート解除予定','負担割合整合'
    ]);

    standardizeConsentNewsMeta_();

    props.setProperty(AUX_SHEETS_INIT_KEY, '1');
  } finally {
    if (locked) {
      lock.releaseLock();
    }
  }
}

function standardizeConsentNewsMeta_(){
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(CONSENT_NEWS_META_STANDARDIZED_KEY) === '1') {
    return;
  }

  let sheet;
  try {
    sheet = sh('News');
  } catch (err) {
    Logger.log('[standardizeConsentNewsMeta_] Failed to open News sheet: ' + (err && err.message ? err.message : err));
    return;
  }

  const lastRow = sheet.getLastRow();
  const width = Math.min(6, sheet.getMaxColumns());
  if (lastRow < 2 || width < 6) {
    props.setProperty(CONSENT_NEWS_META_STANDARDIZED_KEY, '1');
    return;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const updates = [];

  values.forEach((row, idx) => {
    const type = String(row[2] || '').trim();
    if (type !== '同意') return;
    const message = String(row[3] || '');
    const metaRaw = row[5];
    const parsedMeta = parseNewsMetaValue_(metaRaw);
    const normalized = normalizeConsentNewsMeta_(parsedMeta, message);
    if (!normalized.changed) return;
    let cellValue = '';
    if (normalized.meta != null) {
      try {
        cellValue = typeof normalized.meta === 'string' ? normalized.meta : JSON.stringify(normalized.meta);
      } catch (err) {
        cellValue = String(normalized.meta);
      }
    }
    updates.push({ row: 2 + idx, value: cellValue });
  });

  updates.forEach(update => {
    sheet.getRange(update.row, 6).setValue(update.value);
  });

  props.setProperty(CONSENT_NEWS_META_STANDARDIZED_KEY, '1');

  if (updates.length) {
    invalidateGlobalNewsCache_();
  }
}

function ensureAiReportSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName('AI報告書');
  if (!sheet) {
    sheet = wb.insertSheet('AI報告書');
  }
  const needed = AI_REPORT_SHEET_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([AI_REPORT_SHEET_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || AI_REPORT_SHEET_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([AI_REPORT_SHEET_HEADER]);
  }
  return sheet;
}

function ensureVisitAttendanceSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(VISIT_ATTENDANCE_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(VISIT_ATTENDANCE_SHEET_NAME);
  }
  const needed = VISIT_ATTENDANCE_SHEET_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([VISIT_ATTENDANCE_SHEET_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || VISIT_ATTENDANCE_SHEET_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([VISIT_ATTENDANCE_SHEET_HEADER]);
  }
  return sheet;
}

function ensureVisitAttendanceRequestSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(VISIT_ATTENDANCE_REQUEST_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(VISIT_ATTENDANCE_REQUEST_SHEET_NAME);
  }
  const needed = VISIT_ATTENDANCE_REQUEST_SHEET_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([VISIT_ATTENDANCE_REQUEST_SHEET_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || VISIT_ATTENDANCE_REQUEST_SHEET_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([VISIT_ATTENDANCE_REQUEST_SHEET_HEADER]);
  }
  return sheet;
}

function ensureVisitAttendanceStaffSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(VISIT_ATTENDANCE_STAFF_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(VISIT_ATTENDANCE_STAFF_SHEET_NAME);
  }
  const needed = VISIT_ATTENDANCE_STAFF_SHEET_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([VISIT_ATTENDANCE_STAFF_SHEET_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || VISIT_ATTENDANCE_STAFF_SHEET_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([VISIT_ATTENDANCE_STAFF_SHEET_HEADER]);
  }
  return sheet;
}

function ensurePayrollEmployeeSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(PAYROLL_EMPLOYEE_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(PAYROLL_EMPLOYEE_SHEET_NAME);
  }
  const needed = PAYROLL_EMPLOYEE_SHEET_HEADER.length;
  if (sheet.getLastRow() === 0) {
    if (sheet.getMaxColumns() < needed) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
    }
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_EMPLOYEE_SHEET_HEADER]);
    return sheet;
  }
  const currentHeaderWidth = sheet.getLastColumn();
  const currentHeaderValues = currentHeaderWidth > 0
    ? sheet.getRange(1, 1, 1, currentHeaderWidth).getDisplayValues()[0]
    : [];
  const hasMunicipalTaxColumn = currentHeaderValues.some(label => String(label || '').trim() === '住民税');
  if (!hasMunicipalTaxColumn) {
    const insertPosition = Math.min(PAYROLL_EMPLOYEE_COLUMN_INDEX.housingDeduction, sheet.getMaxColumns() || PAYROLL_EMPLOYEE_COLUMN_INDEX.housingDeduction);
    sheet.insertColumnAfter(insertPosition);
  }
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || PAYROLL_EMPLOYEE_SHEET_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_EMPLOYEE_SHEET_HEADER]);
  }
  return sheet;
}

function ensurePayrollGradeSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(PAYROLL_GRADE_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(PAYROLL_GRADE_SHEET_NAME);
  }
  const needed = PAYROLL_GRADE_SHEET_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_GRADE_SHEET_HEADER]);
    seedDefaultPayrollGrades_(sheet);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || PAYROLL_GRADE_SHEET_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_GRADE_SHEET_HEADER]);
  }
  if (sheet.getLastRow() <= 1) {
    seedDefaultPayrollGrades_(sheet);
  }
  return sheet;
}

function ensurePayrollSocialInsuranceStandardSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(PAYROLL_SOCIAL_INSURANCE_STANDARD_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(PAYROLL_SOCIAL_INSURANCE_STANDARD_SHEET_NAME);
  }
  const needed = PAYROLL_SOCIAL_INSURANCE_STANDARD_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_SOCIAL_INSURANCE_STANDARD_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || PAYROLL_SOCIAL_INSURANCE_STANDARD_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_SOCIAL_INSURANCE_STANDARD_HEADER]);
  }
  return sheet;
}

function ensurePayrollSocialInsuranceOverrideSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(PAYROLL_SOCIAL_INSURANCE_OVERRIDE_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(PAYROLL_SOCIAL_INSURANCE_OVERRIDE_SHEET_NAME);
  }
  const needed = PAYROLL_SOCIAL_INSURANCE_OVERRIDE_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_SOCIAL_INSURANCE_OVERRIDE_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || PAYROLL_SOCIAL_INSURANCE_OVERRIDE_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_SOCIAL_INSURANCE_OVERRIDE_HEADER]);
  }
  return sheet;
}

function ensurePayrollRoleSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(PAYROLL_ROLE_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(PAYROLL_ROLE_SHEET_NAME);
  }
  const needed = PAYROLL_ROLE_SHEET_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_ROLE_SHEET_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || PAYROLL_ROLE_SHEET_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_ROLE_SHEET_HEADER]);
  }
  return sheet;
}

function ensurePayrollPayoutEventSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(PAYROLL_PAYOUT_EVENT_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(PAYROLL_PAYOUT_EVENT_SHEET_NAME);
  }
  const needed = PAYROLL_PAYOUT_EVENT_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_PAYOUT_EVENT_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || PAYROLL_PAYOUT_EVENT_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_PAYOUT_EVENT_HEADER]);
  }
  return sheet;
}

function ensurePayrollAnnualSummarySheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(PAYROLL_ANNUAL_SUMMARY_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(PAYROLL_ANNUAL_SUMMARY_SHEET_NAME);
  }
  const needed = PAYROLL_ANNUAL_SUMMARY_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_ANNUAL_SUMMARY_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || PAYROLL_ANNUAL_SUMMARY_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([PAYROLL_ANNUAL_SUMMARY_HEADER]);
  }
  return sheet;
}

function seedDefaultPayrollGrades_(sheet){
  if (!sheet || !PAYROLL_GRADE_DEFAULTS || PAYROLL_GRADE_DEFAULTS.length === 0) return;
  const headerRows = sheet.getLastRow();
  if (headerRows > 1) return;
  const now = new Date();
  const rows = PAYROLL_GRADE_DEFAULTS.map(def => [
    Utilities.getUuid(),
    def && def.name ? def.name : '',
    def && def.amount != null ? def.amount : '',
    def && def.note ? def.note : '',
    now
  ]);
  sheet.getRange(2, 1, rows.length, PAYROLL_GRADE_SHEET_HEADER.length).setValues(rows);
}

/***** アルバイト勤怠：共通ユーティリティ *****/
function normalizeAlbyteName_(name){
  return String(name || '').replace(/\u3000/g, ' ').trim();
}

function ensureAlbyteSessionSecret_(){
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty(ALBYTE_SESSION_SECRET_PROPERTY_KEY);
  if (secret) return secret;

  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    locked = lock.tryLock(3000);
  } catch (err) {
    locked = false;
  }

  try {
    secret = props.getProperty(ALBYTE_SESSION_SECRET_PROPERTY_KEY);
    if (!secret) {
      secret = Utilities.getUuid().replace(/-/g, '');
      props.setProperty(ALBYTE_SESSION_SECRET_PROPERTY_KEY, secret);
    }
  } finally {
    if (locked) {
      lock.releaseLock();
    }
  }
  return secret;
}

function createAlbyteSessionToken_(staffId){
  const issuedAt = Date.now();
  const payload = String(staffId || '') + '.' + issuedAt;
  const secret = ensureAlbyteSessionSecret_();
  const sigBytes = Utilities.computeHmacSha256Signature(payload, secret);
  const signature = Utilities.base64EncodeWebSafe(sigBytes);
  return payload + '.' + signature;
}

function validateAlbyteSessionToken_(token){
  const raw = String(token || '').trim();
  if (!raw) return null;
  const parts = raw.split('.');
  if (parts.length !== 3) return null;
  const staffId = parts[0];
  const issuedAtStr = parts[1];
  const signature = parts[2];
  if (!staffId || !issuedAtStr || !signature) return null;
  const payload = staffId + '.' + issuedAtStr;
  const secret = ensureAlbyteSessionSecret_();
  const expected = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret));
  if (expected !== signature) return null;
  const issuedAt = Number(issuedAtStr);
  if (!isFinite(issuedAt)) return null;
  if (Date.now() - issuedAt > ALBYTE_SESSION_TTL_MILLIS) return null;
  return { staffId, issuedAt };
}

function withAlbyteLock_(callback){
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    locked = lock.tryLock(5000);
    if (!locked) {
      throw new Error('現在システムが混み合っています。数秒後に再度お試しください。');
    }
    return callback();
  } finally {
    if (locked) {
      lock.releaseLock();
    }
  }
}

function wrapAlbyteResponse_(tag, fn){
  try {
    return fn();
  } catch (err) {
    const message = err && err.message ? err.message : String(err);
    Logger.log('[%s] %s', tag, err && err.stack ? err.stack : message);
    return { ok: false, reason: 'system_error', message: message || 'エラーが発生しました。' };
  }
}

function ensureAlbyteStaffSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(ALBYTE_STAFF_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(ALBYTE_STAFF_SHEET_NAME);
  }
  const needed = ALBYTE_STAFF_SHEET_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([ALBYTE_STAFF_SHEET_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || ALBYTE_STAFF_SHEET_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([ALBYTE_STAFF_SHEET_HEADER]);
  }
  return sheet;
}

function ensureAlbyteAttendanceSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(ALBYTE_ATTENDANCE_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(ALBYTE_ATTENDANCE_SHEET_NAME);
  }
  const needed = ALBYTE_ATTENDANCE_SHEET_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([ALBYTE_ATTENDANCE_SHEET_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || ALBYTE_ATTENDANCE_SHEET_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([ALBYTE_ATTENDANCE_SHEET_HEADER]);
  }
  return sheet;
}

function ensureAlbyteShiftSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName(ALBYTE_SHIFT_SHEET_NAME);
  if (!sheet) {
    sheet = wb.insertSheet(ALBYTE_SHIFT_SHEET_NAME);
  }
  const needed = ALBYTE_SHIFT_SHEET_HEADER.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed).setValues([ALBYTE_SHIFT_SHEET_HEADER]);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, needed).getDisplayValues()[0];
  const mismatch = current.length < needed || ALBYTE_SHIFT_SHEET_HEADER.some((label, idx) => String(current[idx] || '') !== label);
  if (mismatch) {
    sheet.getRange(1, 1, 1, needed).setValues([ALBYTE_SHIFT_SHEET_HEADER]);
  }
  return sheet;
}

function parseDateValue_(value){
  if (value instanceof Date) return value;
  if (value == null || value === '') return null;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function escapeHtml_(value){
  const text = value == null ? '' : String(value);
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function convertPlainTextToSafeHtml_(text){
  const escaped = escapeHtml_(text || '');
  return escaped.replace(/\r?\n/g, '<br />');
}

function sanitizeDriveFileName_(value){
  const text = String(value || '').trim();
  if (!text) return 'file';
  return text.replace(/[\\/:*?"<>|#%]/g, '_').replace(/\s+/g, ' ').trim();
}

function formatPayrollCurrencyYen_(value){
  const num = Math.round(Number(value) || 0);
  return '¥' + num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatJapaneseEraMonthLabel_(date, tz){
  const base = date instanceof Date && !isNaN(date.getTime()) ? date : new Date();
  const timezone = tz || Session.getScriptTimeZone() || 'Asia/Tokyo';
  const reiwaStart = new Date(2019, 4, 1);
  if (base.getTime() >= reiwaStart.getTime()) {
    const reiwaYear = base.getFullYear() - 2018;
    return '令和' + reiwaYear + '年' + Utilities.formatDate(base, timezone, 'M') + '月分';
  }
  return Utilities.formatDate(base, timezone, 'yyyy年M月分');
}

function readAlbyteStaffRecords_(){
  const sheet = ensureAlbyteStaffSheet_();
  const lastRow = sheet.getLastRow();
  const width = ALBYTE_STAFF_SHEET_HEADER.length;
  const records = [];
  const mapByName = new Map();
  const mapById = new Map();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const isEmpty = row.every(cell => cell === '' || cell == null);
      if (isEmpty) {
        continue;
      }
      const rowIndex = i + 2;
      let id = String(row[ALBYTE_STAFF_COLUMNS.id] || '').trim();
      if (!id) {
        id = Utilities.getUuid();
        sheet.getRange(rowIndex, ALBYTE_STAFF_COLUMN_INDEX.id).setValue(id);
      }
      const name = String(row[ALBYTE_STAFF_COLUMNS.name] || '').trim();
      const normalizedName = normalizeAlbyteName_(name);
      const pin = String(row[ALBYTE_STAFF_COLUMNS.pin] || '').trim();
      const lockedRaw = row[ALBYTE_STAFF_COLUMNS.locked];
      const locked = lockedRaw === true || String(lockedRaw || '').toLowerCase() === 'true' || String(lockedRaw || '').trim().toUpperCase() === 'LOCKED' || String(lockedRaw || '').trim() === '1';
      const failCount = Number(row[ALBYTE_STAFF_COLUMNS.failCount]) || 0;
      const lastLogin = parseDateValue_(row[ALBYTE_STAFF_COLUMNS.lastLogin]);
      const updatedAt = parseDateValue_(row[ALBYTE_STAFF_COLUMNS.updatedAt]);
      const staffTypeRaw = String(row[ALBYTE_STAFF_COLUMNS.staffType] || '').trim().toLowerCase();
      const staffType = staffTypeRaw === 'daily' ? 'daily' : 'hourly';
      const shiftEndRaw = String(row[ALBYTE_STAFF_COLUMNS.shiftEndTime] || '').trim();
      const shiftEndMinutes = parseTimeTextToMinutes_(shiftEndRaw);
      const normalizedShiftEnd = Number.isFinite(shiftEndMinutes) ? formatMinutesAsTimeText_(shiftEndMinutes) : '';
      const record = {
        rowIndex,
        id,
        name,
        normalizedName,
        pin,
        locked,
        failCount,
        lastLogin,
        updatedAt,
        staffType,
        shiftEndTime: normalizedShiftEnd,
        shiftEndMinutes: Number.isFinite(shiftEndMinutes) ? shiftEndMinutes : NaN
      };
      if (normalizedShiftEnd !== shiftEndRaw) {
        sheet.getRange(rowIndex, ALBYTE_STAFF_COLUMN_INDEX.shiftEndTime).setValue(normalizedShiftEnd);
      }
      records.push(record);
      if (normalizedName) {
        mapByName.set(normalizedName, record);
      }
      if (id) {
        mapById.set(id, record);
      }
    }
  }
  return { sheet, records, mapByName, mapById };
}

function getAlbyteStaffByName_(name){
  const normalized = normalizeAlbyteName_(name);
  if (!normalized) return { sheet: ensureAlbyteStaffSheet_(), record: null };
  const context = readAlbyteStaffRecords_();
  return { sheet: context.sheet, record: context.mapByName.get(normalized) || null };
}

function getAlbyteStaffById_(id){
  const context = readAlbyteStaffRecords_();
  return { sheet: context.sheet, record: context.mapById.get(String(id || '').trim()) || null };
}

function formatTimezoneSuffix_(offset){
  const text = String(offset || '').trim();
  if (!text) return '';
  if (text.length === 5) {
    return text.slice(0, 3) + ':' + text.slice(3);
  }
  return text;
}

function formatIsoStringWithOffset_(date, tz){
  const iso = Utilities.formatDate(date, tz, "yyyy-MM-dd'T'HH:mm:ss");
  const offset = formatTimezoneSuffix_(Utilities.formatDate(date, tz, 'Z'));
  return iso + offset;
}

function getWeekdaySymbol_(date, tz){
  const index = Number(Utilities.formatDate(date, tz, 'u'));
  const map = { 1: '月', 2: '火', 3: '水', 4: '木', 5: '金', 6: '土', 7: '日' };
  return map[index] || '';
}

function formatDisplayDateTime_(date, tz){
  const day = Utilities.formatDate(date, tz, 'yyyy年M月d日');
  const time = Utilities.formatDate(date, tz, 'HH:mm');
  const weekday = getWeekdaySymbol_(date, tz);
  return day + (weekday ? '(' + weekday + ')' : '') + ' ' + time;
}

function parseAlbyteAttendanceLog_(value){
  if (!value) return [];
  if (Array.isArray(value)) return value.slice();
  if (typeof value === 'string') {
    const text = value.trim();
    if (!text) return [];
    try {
      const parsed = JSON.parse(text);
      return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
      Logger.log('[albyte] failed to parse log: %s', err && err.message ? err.message : err);
      return [];
    }
  }
  return [];
}

function serializeAlbyteAttendanceLog_(entries){
  if (!Array.isArray(entries)) return '[]';
  try {
    return JSON.stringify(entries);
  } catch (err) {
    Logger.log('[albyte] failed to serialize log: %s', err && err.message ? err.message : err);
    return '[]';
  }
}

function appendAlbyteAttendanceLog_(existingLog, entry){
  const list = parseAlbyteAttendanceLog_(existingLog);
  list.push(entry);
  return serializeAlbyteAttendanceLog_(list);
}

function normalizeAlbyteClockValue_(value){
  if (value == null || value === '') return null;
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, 'HH:mm');
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    if (value >= 0 && value <= 1) {
      return formatMinutesAsTimeText_(Math.round(value * 24 * 60));
    }
    return formatMinutesAsTimeText_(value);
  }
  const text = String(value).trim();
  if (!text) return null;
  const normalized = text.replace(/[時h]/gi, ':').replace(/分/g, '');
  const match = normalized.match(/^(\d{1,2})(?::?(\d{2}))?$/);
  if (match) {
    const hours = match[1].padStart(2, '0');
    const mins = (match[2] || '0').padStart(2, '0');
    return hours + ':' + mins;
  }
  const parsedDate = new Date(text);
  if (!isNaN(parsedDate.getTime())) {
    return Utilities.formatDate(parsedDate, tz, 'HH:mm');
  }
  return text;
}

function deriveAlbyteDateKeyFromLogEntries_(entries){
  if (!Array.isArray(entries)) return '';
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  for (let i = entries.length - 1; i >= 0; i--) {
    const entry = entries[i];
    if (!entry || !entry.at) continue;
    const parsed = new Date(entry.at);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
    }
  }
  return '';
}

function resolveAlbyteAttendanceRecordDate_(record){
  if (!record) return '';
  let dateKey = normalizeDateKey_(record.date);
  if (!dateKey) {
    dateKey = deriveAlbyteDateKeyFromLogEntries_(record.log);
  }
  if (!dateKey) {
    const tz = getConfig('timezone') || 'Asia/Tokyo';
    const candidate = record.updatedAt || record.createdAt;
    if (candidate instanceof Date && !isNaN(candidate.getTime())) {
      dateKey = Utilities.formatDate(candidate, tz, 'yyyy-MM-dd');
    }
  }
  if (dateKey) {
    record.date = dateKey;
  }
  return dateKey;
}

function parseAlbyteAttendanceRow_(row, rowIndex){
  const log = parseAlbyteAttendanceLog_(row[ALBYTE_ATTENDANCE_COLUMNS.log]);
  const record = {
    rowIndex,
    id: String(row[ALBYTE_ATTENDANCE_COLUMNS.id] || '').trim(),
    staffId: String(row[ALBYTE_ATTENDANCE_COLUMNS.staffId] || '').trim(),
    staffName: String(row[ALBYTE_ATTENDANCE_COLUMNS.staffName] || '').trim(),
    date: normalizeDateKey_(row[ALBYTE_ATTENDANCE_COLUMNS.date]),
    clockIn: normalizeAlbyteClockValue_(row[ALBYTE_ATTENDANCE_COLUMNS.clockIn]),
    clockOut: normalizeAlbyteClockValue_(row[ALBYTE_ATTENDANCE_COLUMNS.clockOut]),
    breakMinutes: Number(row[ALBYTE_ATTENDANCE_COLUMNS.breakMinutes]) || 0,
    note: String(row[ALBYTE_ATTENDANCE_COLUMNS.note] || '').trim(),
    autoFlag: String(row[ALBYTE_ATTENDANCE_COLUMNS.autoFlag] || '').trim(),
    log,
    createdAt: parseDateValue_(row[ALBYTE_ATTENDANCE_COLUMNS.createdAt]),
    updatedAt: parseDateValue_(row[ALBYTE_ATTENDANCE_COLUMNS.updatedAt])
  };
  resolveAlbyteAttendanceRecordDate_(record);
  return record;
}

function readAlbyteAttendanceRowFor_(staffId, dateKey, options){
  const sheet = options && options.sheet ? options.sheet : ensureAlbyteAttendanceSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const width = ALBYTE_ATTENDANCE_SHEET_HEADER.length;
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const targetId = String(staffId || '').trim();
  const targetDate = normalizeDateKey_(dateKey);
  const normalizedStaffName = options && options.normalizedStaffName
    ? String(options.normalizedStaffName).trim()
    : (options && options.staff
      ? (options.staff.normalizedName || normalizeAlbyteName_(options.staff.name))
      : '');
  const allowNameFallback = Boolean(options && options.allowNameFallback && normalizedStaffName);
  let fallback = null;
  let match = null;
  const maybeEnsureId = parsed => {
    if (parsed && !parsed.id) {
      parsed.id = Utilities.getUuid();
      sheet.getRange(parsed.rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.id).setValue(parsed.id);
    }
    return parsed;
  };
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowDate = normalizeDateKey_(row[ALBYTE_ATTENDANCE_COLUMNS.date]);
    if (rowDate !== targetDate) continue;
    const rowStaffId = String(row[ALBYTE_ATTENDANCE_COLUMNS.staffId] || '').trim();
    if (targetId && rowStaffId === targetId) {
      match = maybeEnsureId(parseAlbyteAttendanceRow_(row, i + 2));
      continue;
    }
    if (!allowNameFallback || rowStaffId) continue;
    const rowName = normalizeAlbyteName_(row[ALBYTE_ATTENDANCE_COLUMNS.staffName]);
    if (rowName && rowName === normalizedStaffName) {
      fallback = maybeEnsureId(parseAlbyteAttendanceRow_(row, i + 2));
    }
  }
  return match || fallback;
}

function readLatestAlbyteAttendanceRowForStaff_(staffRecord, options){
  if (!staffRecord) return null;
  const sheet = options && options.sheet ? options.sheet : ensureAlbyteAttendanceSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const width = ALBYTE_ATTENDANCE_SHEET_HEADER.length;
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const targetId = String(staffRecord.id || '').trim();
  const normalizedStaffName = staffRecord.normalizedName || normalizeAlbyteName_(staffRecord.name);
  let latest = null;
  const maybeEnsureId = parsed => {
    if (parsed && !parsed.id) {
      parsed.id = Utilities.getUuid();
      sheet.getRange(parsed.rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.id).setValue(parsed.id);
    }
    return parsed;
  };
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowStaffId = String(row[ALBYTE_ATTENDANCE_COLUMNS.staffId] || '').trim();
    let matches = false;
    if (targetId && rowStaffId === targetId) {
      matches = true;
    } else if (!rowStaffId && normalizedStaffName) {
      const rowName = normalizeAlbyteName_(row[ALBYTE_ATTENDANCE_COLUMNS.staffName]);
      matches = rowName && rowName === normalizedStaffName;
    }
    if (!matches) continue;
    const parsed = maybeEnsureId(parseAlbyteAttendanceRow_(row, i + 2));
    if (!parsed) continue;
    const timestamp = (parsed.updatedAt && parsed.updatedAt.getTime())
      || (parsed.createdAt && parsed.createdAt.getTime())
      || 0;
    if (!latest || timestamp >= latest.timestamp) {
      latest = { record: parsed, timestamp };
    }
  }
  return latest ? latest.record : null;
}

function getAlbyteAttendanceById_(id, options){
  const sheet = options && options.sheet ? options.sheet : ensureAlbyteAttendanceSheet_();
  const target = String(id || '').trim();
  if (!target) return { sheet, record: null };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { sheet, record: null };
  const width = ALBYTE_ATTENDANCE_SHEET_HEADER.length;
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const parsed = parseAlbyteAttendanceRow_(row, i + 2);
    if (parsed && parsed.id === target) {
      return { sheet, record: parsed };
    }
  }
  return { sheet, record: null };
}

function normalizeDateKey_(value, tz){
  if (!value && value !== 0) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz || getConfig('timezone') || 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  if (!text) return '';
  const direct = new Date(text);
  if (!isNaN(direct.getTime())) {
    return Utilities.formatDate(direct, tz || getConfig('timezone') || 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return text;
}

function readAlbyteShiftRecords_(){
  const sheet = ensureAlbyteShiftSheet_();
  const lastRow = sheet.getLastRow();
  const width = ALBYTE_SHIFT_SHEET_HEADER.length;
  const records = [];
  if (lastRow < 2) {
    return { sheet, records };
  }
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const isEmpty = row.every(cell => cell === '' || cell == null);
    if (isEmpty) continue;
    const rowIndex = i + 2;
    let id = String(row[ALBYTE_SHIFT_COLUMNS.id] || '').trim();
    if (!id) {
      id = Utilities.getUuid();
      sheet.getRange(rowIndex, ALBYTE_SHIFT_COLUMN_INDEX.id).setValue(id);
    }
    const dateKey = normalizeDateKey_(row[ALBYTE_SHIFT_COLUMNS.date], tz);
    const staffId = String(row[ALBYTE_SHIFT_COLUMNS.staffId] || '').trim();
    const staffName = String(row[ALBYTE_SHIFT_COLUMNS.staffName] || '').trim();
    const normalizedName = normalizeAlbyteName_(staffName);
    const startMinutes = parseTimeTextToMinutes_(row[ALBYTE_SHIFT_COLUMNS.start]);
    const endMinutes = parseTimeTextToMinutes_(row[ALBYTE_SHIFT_COLUMNS.end]);
    const note = String(row[ALBYTE_SHIFT_COLUMNS.note] || '').trim();
    const updatedAt = parseDateValue_(row[ALBYTE_SHIFT_COLUMNS.updatedAt]);
    const record = {
      rowIndex,
      id,
      dateKey,
      staffId,
      staffName,
      normalizedName,
      startMinutes: Number.isFinite(startMinutes) ? startMinutes : NaN,
      endMinutes: Number.isFinite(endMinutes) ? endMinutes : NaN,
      startText: Number.isFinite(startMinutes) ? formatMinutesAsTimeText_(startMinutes) : '',
      endText: Number.isFinite(endMinutes) ? formatMinutesAsTimeText_(endMinutes) : '',
      note,
      updatedAt
    };
    records.push(record);
  }
  return { sheet, records };
}

function getAlbyteShiftById_(id, context){
  const ctx = context || readAlbyteShiftRecords_();
  const target = String(id || '').trim();
  if (!target) return { sheet: ctx.sheet, record: null };
  for (let i = 0; i < ctx.records.length; i++) {
    const record = ctx.records[i];
    if (record && record.id === target) {
      return { sheet: ctx.sheet, record };
    }
  }
  return { sheet: ctx.sheet, record: null };
}

function findAlbyteShiftFor_(staffRecord, dateKey, options){
  const targetDate = String(dateKey || '').trim();
  if (!targetDate) return null;
  const context = options && options.context ? options.context : readAlbyteShiftRecords_();
  const normalizedName = staffRecord && staffRecord.normalizedName
    ? staffRecord.normalizedName
    : normalizeAlbyteName_(staffRecord && staffRecord.name);
  for (let i = 0; i < context.records.length; i++) {
    const record = context.records[i];
    if (!record || record.dateKey !== targetDate) continue;
    if (staffRecord && record.staffId && staffRecord.id && record.staffId === staffRecord.id) {
      return record;
    }
    if (normalizedName && record.normalizedName && normalizedName === record.normalizedName) {
      return record;
    }
  }
  return null;
}

function resolveStaffShiftEndMinutes_(staff, shift){
  if (staff && Number.isFinite(staff.shiftEndMinutes)) {
    return staff.shiftEndMinutes;
  }
  if (shift && Number.isFinite(shift.endMinutes)) {
    return shift.endMinutes;
  }
  return NaN;
}

function roundUpMinutes_(value, unit){
  if (!Number.isFinite(value) || value <= 0) return 0;
  if (!Number.isFinite(unit) || unit <= 0) return Math.max(0, Math.round(value));
  return Math.ceil(value / unit) * unit;
}

function resolveAlbyteWorkMinutes_(clockInText, clockOutText, breakMinutes){
  const startMinutes = parseTimeTextToMinutes_(clockInText);
  const endMinutes = parseTimeTextToMinutes_(clockOutText);
  const breakValue = Number(breakMinutes) || 0;
  if (!Number.isFinite(startMinutes) || !Number.isFinite(endMinutes)) return NaN;
  return Math.max(0, endMinutes - startMinutes - breakValue);
}

function computeAlbyteWorkMetrics_(record, staff, shift){
  const staffType = staff && staff.staffType ? staff.staffType : 'hourly';
  const isDailyStaff = staffType === 'daily';
  const breakMinutesRaw = Number(record && record.breakMinutes) || 0;
  const breakMinutes = isDailyStaff ? 0 : breakMinutesRaw;
  const clockInText = record && record.clockIn ? record.clockIn : '';
  const clockOutText = record && record.clockOut ? record.clockOut : '';
  const clockInMinutes = parseTimeTextToMinutes_(clockInText);
  const clockOutMinutes = parseTimeTextToMinutes_(clockOutText);
  const baseShiftEndMinutes = resolveStaffShiftEndMinutes_(staff, shift);
  const baseShiftEndText = Number.isFinite(baseShiftEndMinutes) ? formatMinutesAsTimeText_(baseShiftEndMinutes) : '';
  const resolvedWork = resolveAlbyteWorkMinutes_(clockInText, clockOutText, breakMinutes);
  let workMinutes = Number.isFinite(resolvedWork) ? resolvedWork : NaN;
  let overtimeMinutes = 0;
  if (isDailyStaff && Number.isFinite(clockOutMinutes) && Number.isFinite(baseShiftEndMinutes)) {
    const diff = clockOutMinutes - baseShiftEndMinutes;
    if (diff > 0) {
      overtimeMinutes = roundUpMinutes_(diff, ALBYTE_DAILY_OVERTIME_ROUNDING_MINUTES);
    }
    workMinutes = overtimeMinutes;
  }
  return {
    staffType,
    isDailyStaff,
    breakMinutes,
    workMinutes,
    overtimeMinutes,
    baseShiftEndMinutes,
    baseShiftEndText
  };
}

function resolveAlbyteSourceLabel_(record, metrics, shift){
  const fallback = '通常勤務';
  if (!record) return fallback;
  if (metrics && metrics.isDailyStaff) {
    return metrics.overtimeMinutes > 0 ? '延長勤務' : fallback;
  }
  const clockOutMinutes = parseTimeTextToMinutes_(record.clockOut);
  if (shift && Number.isFinite(shift.endMinutes) && Number.isFinite(clockOutMinutes)) {
    if (clockOutMinutes > shift.endMinutes) {
      return '延長勤務';
    }
  }
  return fallback;
}

function applyAlbyteAutoAdjustmentsForRow_(attendanceRecord, options){
  if (!attendanceRecord || !attendanceRecord.rowIndex) {
    return { record: attendanceRecord, autoAdjusted: false, shift: null, workMinutes: resolveAlbyteWorkMinutes_(attendanceRecord && attendanceRecord.clockIn, attendanceRecord && attendanceRecord.clockOut, attendanceRecord && attendanceRecord.breakMinutes) };
  }
  const sheet = options && options.sheet ? options.sheet : ensureAlbyteAttendanceSheet_();
  const staff = options && options.staff ? options.staff : {
    id: attendanceRecord.staffId,
    name: attendanceRecord.staffName,
    normalizedName: normalizeAlbyteName_(attendanceRecord.staffName),
    staffType: 'hourly'
  };
  const shiftContext = options && options.shiftContext ? options.shiftContext : readAlbyteShiftRecords_();
  const shift = findAlbyteShiftFor_(staff, attendanceRecord.date, { context: shiftContext });
  const isDailyStaff = staff && staff.staffType === 'daily';
  let breakMinutes = Number(attendanceRecord.breakMinutes) || 0;

  let clockInMinutes = parseTimeTextToMinutes_(attendanceRecord.clockIn);
  let clockOutMinutes = parseTimeTextToMinutes_(attendanceRecord.clockOut);
  let adjustedClockIn = Number.isFinite(clockInMinutes) ? formatMinutesAsTimeText_(clockInMinutes) : '';
  let adjustedClockOut = Number.isFinite(clockOutMinutes) ? formatMinutesAsTimeText_(clockOutMinutes) : '';
  const messages = [];

  if (shift && Number.isFinite(shift.startMinutes) && Number.isFinite(clockInMinutes)) {
    if (clockInMinutes < shift.startMinutes) {
      clockInMinutes = shift.startMinutes;
      adjustedClockIn = formatMinutesAsTimeText_(shift.startMinutes);
      messages.push('出勤:シフト開始');
    }
  }

  if (shift && Number.isFinite(shift.endMinutes) && Number.isFinite(clockOutMinutes) && !isDailyStaff) {
    if (clockOutMinutes > shift.endMinutes) {
      clockOutMinutes = shift.endMinutes;
      adjustedClockOut = formatMinutesAsTimeText_(shift.endMinutes);
      messages.push('退勤:シフト終了');
    }
  }

  if (isDailyStaff && breakMinutes !== 0) {
    breakMinutes = 0;
    messages.push('休憩:日給スタッフは休憩入力なし');
  }

  const autoFlag = messages.length ? messages.join(' / ') : '';
  const now = new Date();
  let touched = false;
  if (adjustedClockIn !== attendanceRecord.clockIn) {
    sheet.getRange(attendanceRecord.rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.clockIn).setValue(adjustedClockIn);
    touched = true;
  }
  if (adjustedClockOut !== attendanceRecord.clockOut) {
    sheet.getRange(attendanceRecord.rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.clockOut).setValue(adjustedClockOut);
    touched = true;
  }
  if ((attendanceRecord.breakMinutes || 0) !== breakMinutes) {
    sheet.getRange(attendanceRecord.rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.breakMinutes).setValue(breakMinutes);
    touched = true;
  }
  if (autoFlag !== (attendanceRecord.autoFlag || '')) {
    sheet.getRange(attendanceRecord.rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.autoFlag).setValue(autoFlag);
    touched = true;
  }
  if (touched) {
    sheet.getRange(attendanceRecord.rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.updatedAt).setValue(now);
  }

  const updatedRecord = Object.assign({}, attendanceRecord, {
    clockIn: adjustedClockIn,
    clockOut: adjustedClockOut,
    breakMinutes,
    autoFlag,
    updatedAt: touched ? now : attendanceRecord.updatedAt
  });

  const metrics = computeAlbyteWorkMetrics_(updatedRecord, staff, shift);
  const workMinutes = Number.isFinite(metrics.workMinutes) ? metrics.workMinutes : NaN;

  return {
    record: updatedRecord,
    autoAdjusted: messages.length > 0,
    shift,
    workMinutes
  };
}

function readAlbyteAttendanceRecords_(options){
  const sheet = ensureAlbyteAttendanceSheet_();
  const lastRow = sheet.getLastRow();
  const width = ALBYTE_ATTENDANCE_SHEET_HEADER.length;
  const records = [];
  const fromKey = options && options.fromDateKey ? String(options.fromDateKey) : '';
  const toKey = options && options.toDateKey ? String(options.toDateKey) : '';
  const staffId = options && options.staffId ? String(options.staffId) : '';
  const normalizedStaffName = options && options.normalizedStaffName
    ? String(options.normalizedStaffName).trim()
    : '';
  const allowNameFallback = Boolean(options && options.allowNameFallback && normalizedStaffName);
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const parsed = parseAlbyteAttendanceRow_(row, i + 2);
      if (!parsed || !parsed.date) continue;
      if (fromKey && parsed.date < fromKey) continue;
      if (toKey && parsed.date > toKey) continue;
      if (staffId) {
        if (parsed.staffId === staffId) {
          // ok
        } else if (allowNameFallback && !parsed.staffId) {
          const rowName = normalizeAlbyteName_(parsed.staffName);
          if (rowName !== normalizedStaffName) {
            continue;
          }
        } else {
          continue;
        }
      }
      records.push(parsed);
    }
  }
  return { sheet, records };
}

function normalizeYearMonthInput_(year, month){
  const y = Number(year);
  const m = Number(month);
  if (!Number.isFinite(y) || !Number.isFinite(m)) {
    return null;
  }
  if (y <= 0 || m < 1 || m > 12) {
    return null;
  }
  return { year: Math.round(y), month: Math.round(m) };
}

function resolveYearMonthOrCurrent_(year, month){
  const normalized = normalizeYearMonthInput_(year, month);
  if (normalized) {
    return normalized;
  }
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() + 1 };
}

function resolveMonthlyRangeKeys_(year, month){
  const normalized = normalizeYearMonthInput_(year, month);
  if (!normalized) return { from: '', to: '' };
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  const start = new Date(normalized.year, normalized.month - 1, 1);
  const end = new Date(normalized.year, normalized.month, 0);
  return {
    from: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
    to: Utilities.formatDate(end, tz, 'yyyy-MM-dd')
  };
}

function buildAlbyteAttendanceView_(record, options){
  if (!record) return null;
  const shiftContext = options && options.shiftContext ? options.shiftContext : readAlbyteShiftRecords_();
  const staff = options && options.staff ? options.staff : {
    id: record.staffId,
    name: record.staffName,
    normalizedName: normalizeAlbyteName_(record.staffName)
  };
  const shift = findAlbyteShiftFor_(staff, record.date, { context: shiftContext });
  const metrics = computeAlbyteWorkMetrics_(record, staff, shift);
  const breakMinutes = Number.isFinite(metrics.breakMinutes) ? metrics.breakMinutes : 0;
  const workMinutes = Number.isFinite(metrics.workMinutes) ? metrics.workMinutes : NaN;
  const overtimeText = metrics.isDailyStaff ? formatMinutesAsTimeText_(metrics.overtimeMinutes || 0) : '';
  const sourceLabel = resolveAlbyteSourceLabel_(record, metrics, shift);
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  return {
    id: record.id,
    rowIndex: record.rowIndex,
    staffId: record.staffId,
    staffName: record.staffName,
    date: record.date,
    clockIn: record.clockIn || '',
    clockOut: record.clockOut || '',
    breakMinutes,
    breakText: formatMinutesAsTimeText_(breakMinutes),
    workMinutes: Number.isFinite(workMinutes) ? workMinutes : 0,
    workText: Number.isFinite(workMinutes) ? formatMinutesAsTimeText_(workMinutes) : '00:00',
    durationText: Number.isFinite(workMinutes) ? formatDurationText_(workMinutes) : '0時間',
    overtimeText,
    staffType: metrics.staffType,
    isDailyStaff: metrics.isDailyStaff,
    note: record.note || '',
    sourceLabel,
    autoFlag: record.autoFlag || '',
    log: record.log || [],
    createdAt: record.createdAt ? formatIsoStringWithOffset_(record.createdAt, tz) : null,
    updatedAt: record.updatedAt ? formatIsoStringWithOffset_(record.updatedAt, tz) : null,
    shiftStart: shift && shift.startText ? shift.startText : '',
    shiftEnd: shift && shift.endText ? shift.endText : '',
    shiftNote: shift && shift.note ? shift.note : '',
    baseShiftEnd: metrics.baseShiftEndText || (shift && shift.endText ? shift.endText : '')
  };
}

function buildAlbytePortalState_(staffRecord){
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  const now = new Date();
  const todayKey = fmtDate(now, tz);
  const weekday = getWeekdaySymbol_(now, tz);
  let attendance = readAlbyteAttendanceRowFor_(staffRecord.id, todayKey, {
    staff: staffRecord,
    allowNameFallback: true
  });
  if (!attendance) {
    const latest = readLatestAlbyteAttendanceRowForStaff_(staffRecord, {});
    if (latest) {
      const resolvedDate = resolveAlbyteAttendanceRecordDate_(latest);
      if (resolvedDate === todayKey) {
        attendance = latest;
      }
    }
  }
  const staffType = staffRecord && staffRecord.staffType ? staffRecord.staffType : 'hourly';
  const baseShiftEnd = staffRecord && staffRecord.shiftEndTime ? staffRecord.shiftEndTime : '';
  const base = {
    now: {
      iso: formatIsoStringWithOffset_(now, tz),
      display: formatDisplayDateTime_(now, tz)
    },
    today: {
      date: todayKey,
      display: Utilities.formatDate(now, tz, 'yyyy年M月d日') + (weekday ? '(' + weekday + ')' : ''),
      weekday,
      status: 'idle',
      breakMinutes: 0,
      record: null,
      staffType,
      baseShiftEnd
    },
    presets: ALBYTE_BREAK_MINUTES_PRESETS.slice(),
    limits: {
      break: {
        max: ALBYTE_MAX_BREAK_MINUTES,
        step: ALBYTE_BREAK_STEP_MINUTES
      }
    },
    staffType,
    baseShiftEnd
  };
  if (attendance) {
    const hasClockIn = Boolean(attendance.clockIn);
    const hasClockOut = Boolean(attendance.clockOut);
    const status = hasClockIn ? (hasClockOut ? 'completed' : 'working') : 'idle';
    base.today.status = status;
    base.today.breakMinutes = attendance.breakMinutes || 0;
    base.today.record = {
      id: attendance.id,
      clockIn: attendance.clockIn || '',
      clockOut: attendance.clockOut || '',
      breakMinutes: attendance.breakMinutes || 0,
      note: attendance.note || '',
      autoFlag: attendance.autoFlag || '',
      log: attendance.log,
      updatedAt: attendance.updatedAt ? formatIsoStringWithOffset_(attendance.updatedAt, tz) : null,
      createdAt: attendance.createdAt ? formatIsoStringWithOffset_(attendance.createdAt, tz) : null
    };
  }
  return base;
}

function resolveAlbyteSession_(token){
  const parsed = validateAlbyteSessionToken_(token);
  if (!parsed) {
    return { ok: false, reason: 'session_invalid', message: 'セッションが無効です。再度ログインしてください。' };
  }
  const context = readAlbyteStaffRecords_();
  const staff = context.mapById.get(parsed.staffId);
  if (!staff) {
    return { ok: false, reason: 'session_invalid', message: 'セッションが無効です。再度ログインしてください。' };
  }
  if (staff.locked) {
    return { ok: false, reason: 'account_locked', message: 'アカウントがロックされています。管理者に連絡してください。', staff };
  }
  return { ok: true, staff };
}

function buildAlbyteSuccessResponse_(staff, options){
  const portal = buildAlbytePortalState_(staff);
  const response = {
    ok: true,
    staff: {
      id: staff.id,
      name: staff.name,
      staffType: staff.staffType || 'hourly',
      shiftEndTime: staff.shiftEndTime || '',
      isDailyStaff: staff.staffType === 'daily'
    },
    portal
  };
  if (options && options.token) {
    response.token = options.token;
  }
  if (options && options.renewedToken) {
    response.renewedToken = options.renewedToken;
  }
  return response;
}

function albyteLogin(payload){
  return wrapAlbyteResponse_('albyteLogin', () => {
    const nameRaw = payload && payload.name;
    const pinRaw = payload && payload.pin;
    const name = normalizeAlbyteName_(nameRaw);
    const pin = String(pinRaw || '').trim();
    if (!name) {
      return { ok: false, reason: 'validation', message: '名前を入力してください。' };
    }
    if (!/^\d{4}$/.test(pin)) {
      return { ok: false, reason: 'validation', message: 'PINは4桁の数字で入力してください。' };
    }

    return withAlbyteLock_(() => {
      const { sheet, record } = getAlbyteStaffByName_(name);
      if (!record) {
        return { ok: false, reason: 'not_found', message: 'スタッフが見つかりません。管理者に連絡してください。' };
      }
      if (record.locked) {
        return { ok: false, reason: 'account_locked', message: 'アカウントがロックされています。管理者に連絡してください。' };
      }

      const storedPin = String(record.pin || '').trim();
      if (storedPin !== pin) {
        const nextFail = (record.failCount || 0) + 1;
        const willLock = nextFail >= ALBYTE_MAX_PIN_ATTEMPTS;
        const now = new Date();
        sheet.getRange(record.rowIndex, ALBYTE_STAFF_COLUMN_INDEX.locked, 1, 4)
          .setValues([[willLock, nextFail, record.lastLogin || '', now]]);
        record.failCount = nextFail;
        record.locked = willLock;
        record.updatedAt = now;
        return {
          ok: false,
          reason: willLock ? 'account_locked' : 'invalid_pin',
          message: willLock
            ? 'PINを5回連続で間違えたためロックされました。管理者に連絡してください。'
            : 'PINが一致しません。',
          remainingAttempts: willLock ? 0 : Math.max(0, ALBYTE_MAX_PIN_ATTEMPTS - nextFail)
        };
      }

      const now = new Date();
      sheet.getRange(record.rowIndex, ALBYTE_STAFF_COLUMN_INDEX.locked, 1, 4)
        .setValues([[false, 0, now, now]]);
      record.locked = false;
      record.failCount = 0;
      record.lastLogin = now;
      record.updatedAt = now;

      const token = createAlbyteSessionToken_(record.id);
      return buildAlbyteSuccessResponse_(record, { token });
    });
  });
}

function albyteGetPortalState(payload){
  return wrapAlbyteResponse_('albyteGetPortalState', () => {
    const token = payload && payload.token;
    if (!token) {
      return { ok: false, reason: 'session_invalid', message: 'セッションが無効です。再度ログインしてください。' };
    }
    const session = resolveAlbyteSession_(token);
    if (!session.ok) {
      return session;
    }
    return buildAlbyteSuccessResponse_(session.staff, {});
  });
}

function albyteClockIn(payload){
  return wrapAlbyteResponse_('albyteClockIn', () => {
    const token = payload && payload.token;
    if (!token) {
      return { ok: false, reason: 'session_invalid', message: 'セッションが無効です。再度ログインしてください。' };
    }
    const session = resolveAlbyteSession_(token);
    if (!session.ok) {
      return session;
    }

    return withAlbyteLock_(() => {
      const staff = session.staff;
      const tz = getConfig('timezone') || 'Asia/Tokyo';
      const now = new Date();
      const dateKey = fmtDate(now, tz);
      const timeStr = Utilities.formatDate(now, tz, 'HH:mm');
      const iso = formatIsoStringWithOffset_(now, tz);
      const sheet = ensureAlbyteAttendanceSheet_();
      const existing = readAlbyteAttendanceRowFor_(staff.id, dateKey, { sheet });
      if (existing && existing.clockIn) {
        if (existing.clockOut) {
          return { ok: false, reason: 'already_completed', message: '本日の勤怠はすでに退勤済みです。' };
        }
        return { ok: false, reason: 'already_clocked_in', message: 'すでに出勤打刻済みです。' };
      }

      if (!existing) {
        const log = serializeAlbyteAttendanceLog_([{ type: 'clockIn', at: iso }]);
        sheet.appendRow([
          Utilities.getUuid(),
          staff.id,
          staff.name,
          dateKey,
          timeStr,
          '',
          0,
          '',
          '',
          log,
          now,
          now
        ]);
      } else {
        const rowIndex = existing.rowIndex;
        const log = appendAlbyteAttendanceLog_(existing.log, { type: 'clockIn', at: iso });
        sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.staffName).setValue(staff.name);
        sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.clockIn).setValue(timeStr);
        sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.log).setValue(log);
        sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.updatedAt).setValue(now);
        if (!existing.createdAt) {
          sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.createdAt).setValue(now);
        }
      }

      let refreshed = readAlbyteAttendanceRowFor_(staff.id, dateKey, { sheet, staff });
      if (!refreshed) {
        refreshed = readLatestAlbyteAttendanceRowForStaff_(staff, { sheet });
      }
      if (refreshed) {
        applyAlbyteAutoAdjustmentsForRow_(refreshed, { sheet, staff });
      }

      return buildAlbyteSuccessResponse_(staff, {});
    });
  });
}

function albyteClockOut(payload){
  return wrapAlbyteResponse_('albyteClockOut', () => {
    const token = payload && payload.token;
    if (!token) {
      return { ok: false, reason: 'session_invalid', message: 'セッションが無効です。再度ログインしてください。' };
    }
    const session = resolveAlbyteSession_(token);
    if (!session.ok) {
      return session;
    }

    return withAlbyteLock_(() => {
      const staff = session.staff;
      const tz = getConfig('timezone') || 'Asia/Tokyo';
      const now = new Date();
      const dateKey = fmtDate(now, tz);
      const timeStr = Utilities.formatDate(now, tz, 'HH:mm');
      const iso = formatIsoStringWithOffset_(now, tz);
      const sheet = ensureAlbyteAttendanceSheet_();
      const existing = readAlbyteAttendanceRowFor_(staff.id, dateKey, { sheet });
      if (!existing || !existing.clockIn) {
        return { ok: false, reason: 'not_clocked_in', message: '出勤打刻がまだ記録されていません。' };
      }
      if (existing.clockOut) {
        return { ok: false, reason: 'already_clocked_out', message: 'すでに退勤打刻済みです。' };
      }

      const rowIndex = existing.rowIndex;
      const log = appendAlbyteAttendanceLog_(existing.log, { type: 'clockOut', at: iso });
      sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.clockOut).setValue(timeStr);
      sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.log).setValue(log);
      sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.updatedAt).setValue(now);

      let refreshed = readAlbyteAttendanceRowFor_(staff.id, dateKey, { sheet, staff });
      if (!refreshed) {
        refreshed = readLatestAlbyteAttendanceRowForStaff_(staff, { sheet });
      }
      if (refreshed) {
        applyAlbyteAutoAdjustmentsForRow_(refreshed, { sheet, staff });
      }

      return buildAlbyteSuccessResponse_(staff, {});
    });
  });
}

function albyteUpdateBreak(payload){
  return wrapAlbyteResponse_('albyteUpdateBreak', () => {
    const token = payload && payload.token;
    if (!token) {
      return { ok: false, reason: 'session_invalid', message: 'セッションが無効です。再度ログインしてください。' };
    }
    const session = resolveAlbyteSession_(token);
    if (!session.ok) {
      return session;
    }

    if (session.staff && session.staff.staffType === 'daily') {
      return { ok: false, reason: 'not_supported', message: '日給スタッフは休憩登録を利用しません。' };
    }

    const minutesRaw = payload && payload.minutes;
    const minutes = Number(minutesRaw);
    if (!isFinite(minutes)) {
      return { ok: false, reason: 'validation', message: '休憩時間は15分単位で入力してください。' };
    }
    if (minutes < 0) {
      return { ok: false, reason: 'validation', message: '休憩時間は0分以上で入力してください。' };
    }
    if (minutes > ALBYTE_MAX_BREAK_MINUTES) {
      return { ok: false, reason: 'validation', message: '休憩時間は最大180分までです。' };
    }
    if (minutes % ALBYTE_BREAK_STEP_MINUTES !== 0) {
      return { ok: false, reason: 'validation', message: '休憩時間は15分刻みで入力してください。' };
    }

    return withAlbyteLock_(() => {
      const staff = session.staff;
      const tz = getConfig('timezone') || 'Asia/Tokyo';
      const now = new Date();
      const dateKey = fmtDate(now, tz);
      const iso = formatIsoStringWithOffset_(now, tz);
      const sheet = ensureAlbyteAttendanceSheet_();
      const existing = readAlbyteAttendanceRowFor_(staff.id, dateKey, { sheet });
      if (!existing) {
        return { ok: false, reason: 'not_found', message: '本日の勤務データがまだありません。先に出勤打刻を行ってください。' };
      }

      const rowIndex = existing.rowIndex;
      const log = appendAlbyteAttendanceLog_(existing.log, {
        type: 'breakUpdate',
        at: iso,
        minutes,
        source: payload && payload.source ? String(payload.source) : ''
      });
      sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.breakMinutes).setValue(minutes);
      sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.log).setValue(log);
      sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.updatedAt).setValue(now);

      let refreshed = readAlbyteAttendanceRowFor_(staff.id, dateKey, { sheet, staff });
      if (!refreshed) {
        refreshed = readLatestAlbyteAttendanceRowForStaff_(staff, { sheet });
      }
      if (refreshed) {
        applyAlbyteAutoAdjustmentsForRow_(refreshed, { sheet, staff });
      }

      return buildAlbyteSuccessResponse_(staff, {});
    });
  });
}

function albyteGetMonthlySummary(payload){
  return wrapAlbyteResponse_('albyteGetMonthlySummary', () => {
    const token = payload && payload.token;
    if (!token) {
      return { ok: false, reason: 'session_invalid', message: 'セッションが無効です。再度ログインしてください。' };
    }
    const session = resolveAlbyteSession_(token);
    if (!session.ok) {
      return session;
    }

    const staff = session.staff;
    const staffType = staff && staff.staffType ? staff.staffType : 'hourly';
    const isDailyStaff = staffType === 'daily';
    const resolvedMonth = resolveYearMonthOrCurrent_(payload && payload.year, payload && payload.month);
    const year = resolvedMonth.year;
    const month = resolvedMonth.month;
    const { from, to } = resolveMonthlyRangeKeys_(year, month);
    const shiftContext = readAlbyteShiftRecords_();
    const normalizedStaffName = staff && (staff.normalizedName || normalizeAlbyteName_(staff.name));
    const { records } = readAlbyteAttendanceRecords_({
      fromDateKey: from,
      toDateKey: to,
      staffId: staff.id,
      normalizedStaffName,
      allowNameFallback: true
    });
    const list = records.map(record => buildAlbyteAttendanceView_(record, { shiftContext, staff }));
    let totalWork = 0;
    let totalBreak = 0;
    let workingDays = 0;
    list.forEach(item => {
      if (!item) return;
      if (item.workMinutes > 0) {
        workingDays += 1;
      } else if ((item.clockIn && item.clockOut) || item.clockIn || item.clockOut) {
        workingDays += 1;
      }
      totalWork += Number.isFinite(item.workMinutes) ? item.workMinutes : 0;
      totalBreak += Number.isFinite(item.breakMinutes) ? item.breakMinutes : 0;
    });

    const hourlyWage = getAlbyteHourlyWage_();
    const estimatedWage = !isDailyStaff && hourlyWage != null ? Math.round((totalWork / 60) * hourlyWage) : null;

    return {
      ok: true,
      staff: { id: staff.id, name: staff.name, staffType },
      summary: {
        year,
        month,
        range: { from, to },
        records: list,
        totals: {
          workMinutes: totalWork,
          workText: formatMinutesAsTimeText_(totalWork),
          breakMinutes: totalBreak,
          breakText: formatMinutesAsTimeText_(totalBreak),
          workingDays,
          durationText: formatDurationText_(totalWork),
          estimatedWage,
          hourlyWage
        },
        staffType
      }
    };
  });
}

function albyteAdminListStaff(){
  return wrapAlbyteResponse_('albyteAdminListStaff', () => {
    const context = readAlbyteStaffRecords_();
    const tz = getConfig('timezone') || 'Asia/Tokyo';
    const staff = context.records.map(rec => ({
      id: rec.id,
      name: rec.name,
      pin: rec.pin,
      locked: !!rec.locked,
      failCount: rec.failCount || 0,
      lastLogin: rec.lastLogin ? formatIsoStringWithOffset_(rec.lastLogin, tz) : null,
      updatedAt: rec.updatedAt ? formatIsoStringWithOffset_(rec.updatedAt, tz) : null,
      staffType: rec.staffType || 'hourly',
      shiftEndTime: rec.shiftEndTime || ''
    }));
    return { ok: true, staff };
  });
}

function albyteAdminSaveStaff(payload){
  return wrapAlbyteResponse_('albyteAdminSaveStaff', () => {
    const rawName = payload && payload.name;
    const pinRaw = payload && payload.pin;
    const normalized = normalizeAlbyteName_(rawName);
    if (!normalized) {
      return { ok: false, reason: 'validation', message: '名前を入力してください。' };
    }
    const pinText = String(pinRaw != null ? pinRaw : '').trim();
    if (pinText && !/^\d{4}$/.test(pinText)) {
      return { ok: false, reason: 'validation', message: 'PINは4桁の数字で入力してください。' };
    }
    const locked = !!(payload && payload.locked);
    const staffTypeRaw = payload && payload.staffType ? String(payload.staffType).trim().toLowerCase() : '';
    const staffType = staffTypeRaw === 'daily' ? 'daily' : 'hourly';
    const shiftEndRaw = payload && payload.shiftEndTime != null ? String(payload.shiftEndTime).trim() : '';
    const shiftEndMinutes = shiftEndRaw ? parseTimeTextToMinutes_(shiftEndRaw) : NaN;
    let shiftEndText = '';
    if (shiftEndRaw) {
      if (!Number.isFinite(shiftEndMinutes)) {
        return { ok: false, reason: 'validation', message: '基準退勤時刻はHH:MM形式で入力してください。' };
      }
      shiftEndText = formatMinutesAsTimeText_(shiftEndMinutes);
    }
    if (staffType === 'daily' && !shiftEndText) {
      return { ok: false, reason: 'validation', message: '日給スタッフには基準退勤時刻の設定が必要です。' };
    }

    return withAlbyteLock_(() => {
      const context = readAlbyteStaffRecords_();
      const sheet = context.sheet;
      const tz = getConfig('timezone') || 'Asia/Tokyo';
      const now = new Date();
      const idRaw = payload && payload.id;
      let record = idRaw ? context.mapById.get(String(idRaw)) : null;
      let targetId = idRaw ? String(idRaw) : '';

      if (record) {
        const rowIndex = record.rowIndex;
        const failCount = locked ? Math.max(record.failCount || 0, ALBYTE_MAX_PIN_ATTEMPTS) : 0;
        sheet.getRange(rowIndex, ALBYTE_STAFF_COLUMN_INDEX.name, 1, 8)
          .setValues([[rawName, pinText, locked, failCount, record.lastLogin ? record.lastLogin : '', now, staffType, shiftEndText]]);
        record.staffType = staffType;
        record.shiftEndTime = shiftEndText;
        record.shiftEndMinutes = Number.isFinite(shiftEndMinutes) ? shiftEndMinutes : NaN;
        record.pin = pinText;
        record.locked = locked;
        record.failCount = failCount;
        record.updatedAt = now;
      } else {
        const id = Utilities.getUuid();
        const failCount = locked ? ALBYTE_MAX_PIN_ATTEMPTS : 0;
        sheet.appendRow([id, rawName, pinText, locked, failCount, '', now, staffType, shiftEndText]);
        targetId = id;
      }

      const refreshed = readAlbyteStaffRecords_();
      const updated = targetId
        ? refreshed.mapById.get(targetId)
        : refreshed.mapByName.get(normalized);
      if (!updated) {
        return { ok: true, staff: null };
      }
      return {
        ok: true,
        staff: {
          id: updated.id,
          name: updated.name,
          pin: updated.pin,
          locked: !!updated.locked,
          failCount: updated.failCount || 0,
          lastLogin: updated.lastLogin ? formatIsoStringWithOffset_(updated.lastLogin, tz) : null,
          updatedAt: updated.updatedAt ? formatIsoStringWithOffset_(updated.updatedAt, tz) : null,
          staffType: updated.staffType || 'hourly',
          shiftEndTime: updated.shiftEndTime || ''
        }
      };
    });
  });
}

function albyteAdminListAttendance(payload){
  return wrapAlbyteResponse_('albyteAdminListAttendance', () => {
    const resolvedMonth = resolveYearMonthOrCurrent_(payload && payload.year, payload && payload.month);
    const year = resolvedMonth.year;
    const month = resolvedMonth.month;
    const staffId = payload && payload.staffId ? String(payload.staffId).trim() : '';
    const { from, to } = resolveMonthlyRangeKeys_(year, month);
    const staffContext = readAlbyteStaffRecords_();
    const shiftContext = readAlbyteShiftRecords_();
    const { records } = readAlbyteAttendanceRecords_({ fromDateKey: from, toDateKey: to, staffId });
    const list = records.map(record => {
      const staff = staffContext.mapById.get(record.staffId) || {
        id: record.staffId,
        name: record.staffName,
        normalizedName: normalizeAlbyteName_(record.staffName),
        staffType: 'hourly'
      };
      return buildAlbyteAttendanceView_(record, { shiftContext, staff });
    });
    let totalWork = 0;
    let totalBreak = 0;
    list.forEach(item => {
      if (!item) return;
      totalWork += Number.isFinite(item.workMinutes) ? item.workMinutes : 0;
      totalBreak += Number.isFinite(item.breakMinutes) ? item.breakMinutes : 0;
    });
    return {
      ok: true,
      filter: { year, month, from, to, staffId },
      records: list,
      totals: {
        workMinutes: totalWork,
        workText: formatMinutesAsTimeText_(totalWork),
        breakMinutes: totalBreak,
        breakText: formatMinutesAsTimeText_(totalBreak)
      }
    };
  });
}

function albyteAdminSaveAttendance(payload){
  return wrapAlbyteResponse_('albyteAdminSaveAttendance', () => {
    const staffId = String(payload && payload.staffId || '').trim();
    if (!staffId) {
      return { ok: false, reason: 'validation', message: 'スタッフを指定してください。' };
    }
    const tz = getConfig('timezone') || 'Asia/Tokyo';
    const dateKey = normalizeDateKey_(payload && payload.date, tz);
    if (!dateKey) {
      return { ok: false, reason: 'validation', message: '日付を指定してください。' };
    }

    let breakMinutes = Number(payload && payload.breakMinutes);
    if (!Number.isFinite(breakMinutes) || breakMinutes < 0) breakMinutes = 0;
    if (breakMinutes > ALBYTE_MAX_BREAK_MINUTES) {
      return { ok: false, reason: 'validation', message: '休憩時間は最大180分までです。' };
    }
    const startMinutes = parseTimeTextToMinutes_(payload && (payload.clockIn != null ? payload.clockIn : payload.start));
    const endMinutes = parseTimeTextToMinutes_(payload && (payload.clockOut != null ? payload.clockOut : payload.end));
    if (Number.isFinite(startMinutes) && Number.isFinite(endMinutes) && endMinutes < startMinutes) {
      return { ok: false, reason: 'validation', message: '退勤は出勤より後の時刻を指定してください。' };
    }
    const clockInText = Number.isFinite(startMinutes) ? formatMinutesAsTimeText_(startMinutes) : '';
    const clockOutText = Number.isFinite(endMinutes) ? formatMinutesAsTimeText_(endMinutes) : '';
    const noteText = String(payload && payload.note != null ? payload.note : '').trim();

    return withAlbyteLock_(() => {
      const staffContext = readAlbyteStaffRecords_();
      const staff = staffContext.mapById.get(staffId);
      if (!staff) {
        return { ok: false, reason: 'not_found', message: 'スタッフが見つかりません。' };
      }

      const sheet = ensureAlbyteAttendanceSheet_();
      const now = new Date();
      let target;
      if (payload && payload.id) {
        target = getAlbyteAttendanceById_(payload.id, { sheet }).record;
      }
      if (!target) {
        target = readAlbyteAttendanceRowFor_(staffId, dateKey, { sheet });
      }

      if (staff && staff.staffType === 'daily') {
        breakMinutes = 0;
      }

      const actor = (Session.getEffectiveUser && Session.getEffectiveUser().getEmail && Session.getEffectiveUser().getEmail()) || '';
      const logEntry = {
        type: target ? 'adminUpdate' : 'adminCreate',
        at: formatIsoStringWithOffset_(now, tz),
        by: actor,
        payload: {
          clockIn: clockInText,
          clockOut: clockOutText,
          breakMinutes,
          note: noteText
        }
      };

      if (target) {
        const rowIndex = target.rowIndex;
        const log = appendAlbyteAttendanceLog_(target.log, logEntry);
        sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.staffId, 1, 7)
          .setValues([[staff.id, staff.name, dateKey, clockInText, clockOutText, breakMinutes, noteText]]);
        sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.autoFlag).setValue('');
        sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.log).setValue(log);
        if (!target.createdAt) {
          sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.createdAt).setValue(now);
        }
        sheet.getRange(rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.updatedAt).setValue(now);
      } else {
        const id = Utilities.getUuid();
        const log = serializeAlbyteAttendanceLog_([logEntry]);
        sheet.appendRow([
          id,
          staff.id,
          staff.name,
          dateKey,
          clockInText,
          clockOutText,
          breakMinutes,
          noteText,
          '',
          log,
          now,
          now
        ]);
      }

      const refreshed = readAlbyteAttendanceRowFor_(staff.id, dateKey, { sheet });
      if (refreshed) {
        applyAlbyteAutoAdjustmentsForRow_(refreshed, { sheet, staff });
      }
      const updated = readAlbyteAttendanceRowFor_(staff.id, dateKey, { sheet });
      const shiftContext = readAlbyteShiftRecords_();
      const view = updated ? buildAlbyteAttendanceView_(updated, { shiftContext, staff }) : null;

      return { ok: true, record: view };
    });
  });
}

function albyteAdminListShifts(payload){
  return wrapAlbyteResponse_('albyteAdminListShifts', () => {
    const resolvedMonth = resolveYearMonthOrCurrent_(payload && payload.year, payload && payload.month);
    const year = resolvedMonth.year;
    const month = resolvedMonth.month;
    const { from, to } = resolveMonthlyRangeKeys_(year, month);
    const staffContext = readAlbyteStaffRecords_();
    const context = readAlbyteShiftRecords_();
    const tz = getConfig('timezone') || 'Asia/Tokyo';
    const staffId = payload && payload.staffId ? String(payload.staffId).trim() : '';
    const shifts = context.records
      .filter(record => (!from || record.dateKey >= from) && (!to || record.dateKey <= to) && (!staffId || record.staffId === staffId))
      .map(record => ({
        id: record.id,
        date: record.dateKey,
        staffId: record.staffId,
        staffName: record.staffName || (staffContext.mapById.get(record.staffId) ? staffContext.mapById.get(record.staffId).name : ''),
        start: record.startText || '',
        end: record.endText || '',
        note: record.note || '',
        updatedAt: record.updatedAt ? formatIsoStringWithOffset_(record.updatedAt, tz) : null
      }));
    return { ok: true, filter: { year, month, from, to, staffId }, shifts };
  });
}

function albyteAdminSaveShift(payload){
  return wrapAlbyteResponse_('albyteAdminSaveShift', () => {
    const tz = getConfig('timezone') || 'Asia/Tokyo';
    const dateKey = normalizeDateKey_(payload && payload.date, tz);
    if (!dateKey) {
      return { ok: false, reason: 'validation', message: '日付を指定してください。' };
    }
    const staffId = payload && payload.staffId ? String(payload.staffId).trim() : '';
    const staffContext = readAlbyteStaffRecords_();
    let staff = staffId ? staffContext.mapById.get(staffId) : null;
    let staffName = payload && payload.staffName ? String(payload.staffName).trim() : '';
    if (!staff && staffName) {
      staff = staffContext.mapByName.get(normalizeAlbyteName_(staffName)) || null;
    }
    if (staff) {
      staffName = staff.name;
    }
    if (!staff && !staffName) {
      return { ok: false, reason: 'validation', message: 'スタッフを指定してください。' };
    }

    const startMinutes = parseTimeTextToMinutes_(payload && (payload.start != null ? payload.start : payload.shiftStart));
    const endMinutes = parseTimeTextToMinutes_(payload && (payload.end != null ? payload.end : payload.shiftEnd));
    if (Number.isFinite(startMinutes) && Number.isFinite(endMinutes) && endMinutes <= startMinutes) {
      return { ok: false, reason: 'validation', message: 'シフト終了はシフト開始より後の時刻を指定してください。' };
    }
    const startText = Number.isFinite(startMinutes) ? formatMinutesAsTimeText_(startMinutes) : '';
    const endText = Number.isFinite(endMinutes) ? formatMinutesAsTimeText_(endMinutes) : '';
    const noteText = String(payload && payload.note != null ? payload.note : '').trim();

    return withAlbyteLock_(() => {
      const context = readAlbyteShiftRecords_();
      const sheet = context.sheet;
      const now = new Date();
      let entry;
      if (payload && payload.id) {
        entry = getAlbyteShiftById_(payload.id, context).record;
      }
      if (!entry && staff) {
        entry = findAlbyteShiftFor_(staff, dateKey, { context });
      }

      if (entry) {
        const rowIndex = entry.rowIndex;
        sheet.getRange(rowIndex, ALBYTE_SHIFT_COLUMN_INDEX.date, 1, 6)
          .setValues([[dateKey, staff ? staff.id : entry.staffId, staffName || entry.staffName, startText, endText, noteText]]);
        sheet.getRange(rowIndex, ALBYTE_SHIFT_COLUMN_INDEX.updatedAt).setValue(now);
      } else {
        const id = Utilities.getUuid();
        sheet.appendRow([
          id,
          dateKey,
          staff ? staff.id : '',
          staffName,
          startText,
          endText,
          noteText,
          now
        ]);
      }

      const refreshedContext = readAlbyteShiftRecords_();
      const saved = payload && payload.id
        ? getAlbyteShiftById_(payload.id, refreshedContext).record
        : findAlbyteShiftFor_(staff || { id: staffId, name: staffName, normalizedName: normalizeAlbyteName_(staffName) }, dateKey, { context: refreshedContext });

      const staffForAuto = staff
        || (saved && saved.staffId ? staffContext.mapById.get(saved.staffId) : null)
        || (entry && entry.staffId ? staffContext.mapById.get(entry.staffId) : null);
      if (staffForAuto) {
        const attendance = readAlbyteAttendanceRowFor_(staffForAuto.id, dateKey);
        if (attendance) {
          applyAlbyteAutoAdjustmentsForRow_(attendance, { sheet: ensureAlbyteAttendanceSheet_(), staff: staffForAuto, shiftContext: refreshedContext });
        }
      }

      return {
        ok: true,
        shift: saved ? {
          id: saved.id,
          date: saved.dateKey,
          staffId: saved.staffId,
          staffName: saved.staffName,
          start: saved.startText,
          end: saved.endText,
          note: saved.note,
          updatedAt: saved.updatedAt ? formatIsoStringWithOffset_(saved.updatedAt, tz) : null
        } : null
      };
    });
  });
}

function albyteAdminDeleteShift(payload){
  return wrapAlbyteResponse_('albyteAdminDeleteShift', () => {
    const id = payload && payload.id;
    if (!id) {
      return { ok: false, reason: 'validation', message: '削除するシフトを指定してください。' };
    }
    return withAlbyteLock_(() => {
      const context = readAlbyteShiftRecords_();
      const { sheet, record } = getAlbyteShiftById_(id, context);
      if (!record) {
        return { ok: false, reason: 'not_found', message: '対象のシフトが見つかりません。' };
      }
      sheet.deleteRow(record.rowIndex);

      const refreshedContext = readAlbyteShiftRecords_();
      if (record.staffId) {
        const staffContext = readAlbyteStaffRecords_();
        const staff = staffContext.mapById.get(record.staffId);
        if (staff) {
          const attendance = readAlbyteAttendanceRowFor_(staff.id, record.dateKey);
          if (attendance) {
            applyAlbyteAutoAdjustmentsForRow_(attendance, { sheet: ensureAlbyteAttendanceSheet_(), staff, shiftContext: refreshedContext });
          }
        }
      }

      return { ok: true };
    });
  });
}

function albyteGetMonthlyReport(payload){
  return wrapAlbyteResponse_('albyteGetMonthlyReport', () => {
    const resolvedMonth = resolveYearMonthOrCurrent_(payload && payload.year, payload && payload.month);
    const year = resolvedMonth.year;
    const month = resolvedMonth.month;
    const { from, to } = resolveMonthlyRangeKeys_(year, month);
    const staffContext = readAlbyteStaffRecords_();
    const shiftContext = readAlbyteShiftRecords_();
    const { records } = readAlbyteAttendanceRecords_({ fromDateKey: from, toDateKey: to });
    const hourlyWage = getAlbyteHourlyWage_();

    const summaryMap = new Map();
    records.forEach(record => {
      const staff = staffContext.mapById.get(record.staffId) || {
        id: record.staffId,
        name: record.staffName,
        normalizedName: normalizeAlbyteName_(record.staffName)
      };
      const view = buildAlbyteAttendanceView_(record, { shiftContext, staff });
      if (!view) return;
      const key = staff.id || staff.name;
      if (!summaryMap.has(key)) {
        summaryMap.set(key, {
          staffId: staff.id,
          staffName: staff.name,
          workMinutes: 0,
          breakMinutes: 0,
          workingDays: 0,
          records: 0
        });
      }
      const entry = summaryMap.get(key);
      entry.workMinutes += Number.isFinite(view.workMinutes) ? view.workMinutes : 0;
      entry.breakMinutes += Number.isFinite(view.breakMinutes) ? view.breakMinutes : 0;
      if ((view.clockIn && view.clockOut) || view.workMinutes > 0) {
        entry.workingDays += 1;
      }
      entry.records += 1;
    });

    const tz = getConfig('timezone') || 'Asia/Tokyo';
    const rows = Array.from(summaryMap.values()).map(entry => ({
      staffId: entry.staffId,
      staffName: entry.staffName,
      workMinutes: entry.workMinutes,
      workText: formatMinutesAsTimeText_(entry.workMinutes),
      breakMinutes: entry.breakMinutes,
      breakText: formatMinutesAsTimeText_(entry.breakMinutes),
      workingDays: entry.workingDays,
      records: entry.records,
      durationText: formatDurationText_(entry.workMinutes),
      estimatedWage: hourlyWage != null ? Math.round((entry.workMinutes / 60) * hourlyWage) : null
    }));

    let totalWork = 0;
    let totalBreak = 0;
    rows.forEach(entry => {
      totalWork += Number.isFinite(entry.workMinutes) ? entry.workMinutes : 0;
      totalBreak += Number.isFinite(entry.breakMinutes) ? entry.breakMinutes : 0;
    });

    return {
      ok: true,
      report: {
        year,
        month,
        range: { from, to },
        hourlyWage,
        staff: rows,
        totals: {
          workMinutes: totalWork,
          workText: formatMinutesAsTimeText_(totalWork),
          breakMinutes: totalBreak,
          breakText: formatMinutesAsTimeText_(totalBreak),
          durationText: formatDurationText_(totalWork),
          estimatedWage: hourlyWage != null ? Math.round((totalWork / 60) * hourlyWage) : null
        }
      }
    };
  });
}

/*** 給与マスタ（従業員基本情報） ***/
function wrapPayrollResponse_(tag, fn){
  try {
    return fn();
  } catch (err) {
    const message = err && err.message ? err.message : String(err);
    Logger.log('[%s] %s', tag, err && err.stack ? err.stack : message);
    return { ok: false, reason: 'system_error', message: message || 'エラーが発生しました。' };
  }
}

function normalizePayrollGradeName_(value){
  return String(value || '').replace(/\u3000/g, ' ').trim().toLowerCase();
}

function normalizePayrollBaseKey_(value){
  return String(value || '')
    .replace(/\u3000/g, ' ')
    .trim()
    .toLowerCase();
}

function normalizePayrollRoleKey_(value){
  const text = String(value || '').trim().toLowerCase();
  if (text === 'rep' || text === '代表' || text === 'owner' || text === 'ceo') {
    return 'representative';
  }
  if (text === 'manager' || text === 'admin' || text === '管理者') {
    return 'manager';
  }
  return text || '';
}

function normalizePayrollEmploymentType_(value){
  const text = String(value || '').trim().toLowerCase();
  if (!text) return 'employee';
  if (text === 'employee' || text === '正社員' || text === '正社' || text === '社員') {
    return 'employee';
  }
  if (text === 'parttime' || text === 'part-time' || text === 'part_time' || text === 'parttimeemployee' || text === 'アルバイト' || text === 'ﾊﾞｲﾄ' || text === 'パート' || text === 'パートタイム') {
    return 'parttime';
  }
  if (text === 'contractor' || text === '業務委託' || text === '委託' || text === '外注') {
    return 'contractor';
  }
  return 'employee';
}

function formatPayrollEmploymentLabel_(type){
  const key = String(type || '').toLowerCase();
  return PAYROLL_EMPLOYMENT_LABELS[key] || PAYROLL_EMPLOYMENT_LABELS.employee;
}

function normalizePayrollTransportationType_(value){
  const text = String(value || '').trim().toLowerCase();
  if (!text || text === 'none' || text === 'なし' || text === '無' || text === '0') {
    return 'none';
  }
  if (text === 'actual' || text === '実費' || text === 'じっぴ' || text === '実費精算') {
    return 'actual';
  }
  return 'fixed';
}

function formatPayrollTransportationLabel_(type){
  const key = String(type || '').toLowerCase();
  return PAYROLL_TRANSPORTATION_LABELS[key] || PAYROLL_TRANSPORTATION_LABELS.none;
}

function normalizePayrollCommissionLogicType_(value){
  const text = String(value || '').trim().toLowerCase();
  if (text === 'horiguchi' || text === '堀口' || text === '堀口以降') {
    return 'horiguchi';
  }
  return 'legacy';
}

function formatPayrollCommissionLabel_(type){
  const key = String(type || '').toLowerCase();
  return PAYROLL_COMMISSION_LABELS[key] || PAYROLL_COMMISSION_LABELS.legacy;
}

function normalizePayrollPayoutType_(value){
  const text = String(value || '').trim().toLowerCase();
  if (!text) return 'salary';
  if (text === 'bonus' || text === 'shoyo' || text === '賞与') {
    return 'bonus';
  }
  if (text === 'yearend' || text === 'year_end' || text === 'year-end' || text === 'nenmatsu' || text === '年末調整') {
    return 'yearEndAdjustment';
  }
  if (text === 'withholding' || text === 'gensen' || text === '源泉徴収票' || text === 'withholdingcertificate') {
    return 'withholdingCertificate';
  }
  return 'salary';
}

function formatPayrollPayoutLabel_(type){
  const key = String(type || '').trim();
  return PAYROLL_PAYOUT_TYPE_LABELS[key] || PAYROLL_PAYOUT_TYPE_LABELS.salary;
}

function resolvePayrollCommissionMonthRange_(input){
  const now = new Date();
  let year = now.getFullYear();
  let month = now.getMonth() + 1;
  const text = String(input || '').trim();
  if (text) {
    const match = text.match(/^(\d{4})[-/](\d{1,2})$/);
    if (match) {
      const parsedYear = Number(match[1]);
      const parsedMonth = Number(match[2]);
      if (Number.isFinite(parsedYear) && Number.isFinite(parsedMonth) && parsedMonth >= 1 && parsedMonth <= 12) {
        year = parsedYear;
        month = parsedMonth;
      }
    }
  }
  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 1);
  return { year, month, start, end };
}

function collectTreatmentCountsByStaffForRange_(startDate, endDate, tz){
  if (!(startDate instanceof Date) || isNaN(startDate.getTime()) || !(endDate instanceof Date) || isNaN(endDate.getTime())) {
    throw new Error('歩合計算期間が不正です。');
  }
  const startMs = startDate.getTime();
  const endMs = endDate.getTime();
  if (!(endMs > startMs)) {
    throw new Error('歩合計算期間の指定が不正です。');
  }
  const sheet = sh('施術録');
  const lastRow = sheet.getLastRow();
  const width = Math.min(12, sheet.getMaxColumns());
  const results = new Map();
  if (lastRow < 2) {
    return results;
  }
  const rows = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const format = 'yyyy-MM-dd';
  const timezone = tz || Session.getScriptTimeZone() || 'Asia/Tokyo';
  rows.forEach(row => {
    const ts = parseDateValue_(row[0]);
    if (!ts) return;
    const whenMs = ts.getTime();
    if (whenMs < startMs || whenMs >= endMs) return;
    const email = normalizeEmailKey_(row[3]);
    if (!email) return;
    const categoryLabel = width >= 8 ? String(row[7] || '').trim() : '';
    if (!categoryLabel) return;
    let entry = results.get(email);
    if (!entry) {
      entry = { totalCount: 0, byDate: new Map() };
      results.set(email, entry);
    }
    entry.totalCount += 1;
    const dateKey = Utilities.formatDate(ts, timezone, format);
    let dayEntry = entry.byDate.get(dateKey);
    if (!dayEntry) {
      const dateValue = new Date(ts.getFullYear(), ts.getMonth(), ts.getDate());
      dayEntry = { count: 0, dateValue };
      entry.byDate.set(dateKey, dayEntry);
    }
    dayEntry.count += 1;
  });
  return results;
}

function buildPayrollMonthWeekRanges_(startDate, endDate){
  const ranges = [];
  if (!(startDate instanceof Date) || isNaN(startDate.getTime()) || !(endDate instanceof Date) || isNaN(endDate.getTime())) {
    return ranges;
  }
  let cursor = new Date(startDate.getTime());
  while (cursor < endDate) {
    const rangeStart = new Date(cursor.getTime());
    const day = rangeStart.getDay();
    const normalizedDay = day === 0 ? 7 : day;
    const daysUntilNextMonday = 8 - normalizedDay;
    const rangeEnd = new Date(rangeStart.getTime());
    rangeEnd.setDate(rangeEnd.getDate() + daysUntilNextMonday);
    if (rangeEnd > endDate) {
      rangeEnd.setTime(endDate.getTime());
    }
    ranges.push({ start: rangeStart, end: new Date(rangeEnd.getTime()) });
    cursor = rangeEnd;
  }
  return ranges;
}

function sumTreatmentCountsInRange_(entry, rangeStart, rangeEnd){
  if (!entry || !entry.byDate || !(rangeStart instanceof Date) || !(rangeEnd instanceof Date)) {
    return 0;
  }
  let total = 0;
  entry.byDate.forEach(dayEntry => {
    if (!dayEntry || !(dayEntry.dateValue instanceof Date)) return;
    if (dayEntry.dateValue >= rangeStart && dayEntry.dateValue < rangeEnd) {
      const value = Number(dayEntry.count) || 0;
      total += value;
    }
  });
  return total;
}

function buildPayrollCommissionBreakdown_(record, countsEntry, options){
  const tz = (options && options.tz) || Session.getScriptTimeZone() || 'Asia/Tokyo';
  const weekRanges = options && Array.isArray(options.weekRanges) ? options.weekRanges : [];
  const logic = normalizePayrollCommissionLogicType_(record && record.commissionLogic);
  const rules = PAYROLL_COMMISSION_RULES[logic] || PAYROLL_COMMISSION_RULES.legacy;
  const totalTreatments = countsEntry ? Number(countsEntry.totalCount) || 0 : 0;
  const base = {
    id: record && record.id || '',
    name: record && record.name || '',
    email: record && record.email || '',
    commissionLogic: logic,
    commissionLabel: formatPayrollCommissionLabel_(logic),
    totalTreatments
  };
  if (logic === 'horiguchi') {
    const threshold = Number(rules.weeklyThreshold) || 0;
    const weeklyDetails = weekRanges.map(range => {
      const count = sumTreatmentCountsInRange_(countsEntry, range.start, range.end);
      const achieved = threshold > 0 ? count >= threshold : false;
      const endInclusive = new Date(range.end.getTime() - 1);
      return {
        startDate: Utilities.formatDate(range.start, tz, 'yyyy-MM-dd'),
        endDate: Utilities.formatDate(endInclusive, tz, 'yyyy-MM-dd'),
        count,
        achieved
      };
    });
    const achievedWeeks = weeklyDetails.filter(detail => detail.achieved).length;
    const amountPerWeek = Number(rules.amount) || 0;
    return {
      ...base,
      commissionAmount: achievedWeeks * amountPerWeek,
      breakdown: {
        type: 'weekly',
        weeklyThreshold: threshold,
        achievedWeeks,
        weeks: weeklyDetails
      }
    };
  }
  const monthlyThreshold = Number(rules.monthlyThreshold) || 0;
  const achieved = monthlyThreshold > 0 ? totalTreatments >= monthlyThreshold : false;
  return {
    ...base,
    commissionAmount: achieved ? (Number(rules.amount) || 0) : 0,
    breakdown: {
      type: 'monthly',
      monthlyThreshold,
      achieved
    }
  };
}

function normalizePayrollWithholdingType_(value){
  const text = String(value || '').trim().toLowerCase();
  const plain = text.replace(/[（(].*?[）)]/g, '').trim();
  const normalized = plain || text;
  if (!normalized || normalized === 'none' || normalized === '0' || normalized === 'なし' || normalized === '無' || normalized === 'off') {
    return 'none';
  }
  if (normalized === 'required' || normalized === 'あり' || normalized === '有' || normalized === 'true' || normalized === 'yes' || normalized === 'on' || normalized === '1' || normalized === 'withholding') {
    return 'required';
  }
  return 'none';
}

function normalizePayrollWithholdingCategory_(value){
  const text = String(value || '').trim().toLowerCase();
  if (!text || text === 'ko' || text === '甲' || text === '甲欄') {
    return 'ko';
  }
  if (text === '乙' || text === 'otsu') {
    return 'otsu';
  }
  return 'ko';
}

function normalizePayrollWithholdingPeriodType_(value){
  const text = String(value || '').trim().toLowerCase();
  if (text === 'daily' || text === '日額' || text === '日額扱い') {
    return 'daily';
  }
  return 'monthly';
}

function normalizePayrollDependentCount_(value){
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  const clamped = Math.max(0, Math.min(7, Math.floor(num)));
  return clamped;
}

function formatPayrollWithholdingLabel_(type){
  const key = String(type || '').toLowerCase();
  return PAYROLL_WITHHOLDING_LABELS[key] || PAYROLL_WITHHOLDING_LABELS.none;
}

function formatPayrollWithholdingCategoryLabel_(category){
  const key = String(category || '').toLowerCase();
  return PAYROLL_WITHHOLDING_CATEGORY_LABELS[key] || PAYROLL_WITHHOLDING_CATEGORY_LABELS.ko;
}

function formatPayrollWithholdingPeriodLabel_(periodType){
  const key = String(periodType || '').toLowerCase();
  return PAYROLL_WITHHOLDING_PERIOD_LABELS[key] || PAYROLL_WITHHOLDING_PERIOD_LABELS.monthly;
}

function parsePayrollMoneyValue_(value){
  if (value == null || value === '') return null;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return null;
    return Math.round(value);
  }
  const text = String(value || '').trim();
  if (!text) return null;
  const normalized = Number(text.replace(/[^0-9.-]/g, ''));
  if (!Number.isFinite(normalized)) return null;
  return Math.round(normalized);
}

function parseJsonColumnValue_(value){
  if (value == null || value === '') return null;
  if (typeof value === 'object') {
    return value;
  }
  const text = String(value || '').trim();
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch (err) {
    Logger.log('[parseJsonColumnValue_] Failed to parse JSON: ' + (err && err.message ? err.message : err));
    return null;
  }
}

function normalizePayrollMonthKey_(value){
  if (value instanceof Date && !isNaN(value.getTime())) {
    const tz = getConfig('timezone') || 'Asia/Tokyo';
    return Utilities.formatDate(value, tz, 'yyyy-MM');
  }
  const text = String(value || '').trim();
  if (!text) return '';
  const dashMatch = text.match(/^(\d{4})[-/](\d{1,2})$/);
  if (dashMatch) {
    const year = Number(dashMatch[1]);
    const month = Number(dashMatch[2]);
    if (Number.isFinite(year) && Number.isFinite(month) && month >= 1 && month <= 12) {
      return `${year}-${String(month).padStart(2, '0')}`;
    }
  }
  const compactMatch = text.match(/^(\d{4})(\d{2})$/);
  if (compactMatch) {
    const year = Number(compactMatch[1]);
    const month = Number(compactMatch[2]);
    if (Number.isFinite(year) && Number.isFinite(month) && month >= 1 && month <= 12) {
      return `${year}-${String(month).padStart(2, '0')}`;
    }
  }
  const jpMatch = text.match(/^(\d{4})年(\d{1,2})月$/);
  if (jpMatch) {
    const year = Number(jpMatch[1]);
    const month = Number(jpMatch[2]);
    if (Number.isFinite(year) && Number.isFinite(month) && month >= 1 && month <= 12) {
      return `${year}-${String(month).padStart(2, '0')}`;
    }
  }
  return '';
}

function readPayrollGradeRecords_(){
  const sheet = ensurePayrollGradeSheet_();
  const lastRow = sheet.getLastRow();
  const width = PAYROLL_GRADE_SHEET_HEADER.length;
  const records = [];
  const mapById = new Map();
  const mapByNormalizedName = new Map();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    values.forEach((row, idx) => {
      const isEmpty = row.every(cell => cell === '' || cell == null);
      if (isEmpty) return;
      let id = String(row[PAYROLL_GRADE_COLUMNS.id] || '').trim();
      if (!id) {
        id = Utilities.getUuid();
        sheet.getRange(idx + 2, PAYROLL_GRADE_COLUMN_INDEX.id).setValue(id);
      }
      const name = String(row[PAYROLL_GRADE_COLUMNS.name] || '').trim();
      const normalizedName = normalizePayrollGradeName_(name);
      const amount = parsePayrollMoneyValue_(row[PAYROLL_GRADE_COLUMNS.amount]);
      const note = String(row[PAYROLL_GRADE_COLUMNS.note] || '').trim();
      const updatedAt = row[PAYROLL_GRADE_COLUMNS.updatedAt];
      const record = {
        id,
        name,
        normalizedName,
        amount,
        note,
        updatedAt,
        rowIndex: idx + 2
      };
      records.push(record);
      mapById.set(id, record);
      if (normalizedName) {
        mapByNormalizedName.set(normalizedName, record);
      }
    });
  }
  return { sheet, records, mapById, mapByNormalizedName };
}

function readPayrollEmployeeRecords_(){
  const sheet = ensurePayrollEmployeeSheet_();
  const lastRow = sheet.getLastRow();
  const width = PAYROLL_EMPLOYEE_SHEET_HEADER.length;
  const records = [];
  const mapById = new Map();
  const mapByEmail = new Map();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    values.forEach((row, idx) => {
      const isEmpty = row.every(cell => cell === '' || cell == null);
      if (isEmpty) return;
      const rowIndex = idx + 2;
      let id = String(row[PAYROLL_EMPLOYEE_COLUMNS.id] || '').trim();
      if (!id) {
        id = Utilities.getUuid();
        sheet.getRange(rowIndex, PAYROLL_EMPLOYEE_COLUMN_INDEX.id).setValue(id);
      }
      const name = String(row[PAYROLL_EMPLOYEE_COLUMNS.name] || '').trim();
      const email = String(row[PAYROLL_EMPLOYEE_COLUMNS.email] || '').trim();
      const normalizedEmail = normalizeEmailKey_(email);
      const base = String(row[PAYROLL_EMPLOYEE_COLUMNS.base] || '').trim();
      const baseKey = normalizePayrollBaseKey_(base);
      const employmentType = normalizePayrollEmploymentType_(row[PAYROLL_EMPLOYEE_COLUMNS.employmentType]);
      const employmentLabel = formatPayrollEmploymentLabel_(employmentType);
      const baseSalary = parsePayrollMoneyValue_(row[PAYROLL_EMPLOYEE_COLUMNS.baseSalary]);
      const hourlyWage = parsePayrollMoneyValue_(row[PAYROLL_EMPLOYEE_COLUMNS.hourlyWage]);
      const personalAllowance = parsePayrollMoneyValue_(row[PAYROLL_EMPLOYEE_COLUMNS.personalAllowance]);
      const grade = String(row[PAYROLL_EMPLOYEE_COLUMNS.grade] || '').trim();
      const qualificationAllowance = parsePayrollMoneyValue_(row[PAYROLL_EMPLOYEE_COLUMNS.qualificationAllowance]);
      const vehicleAllowance = parsePayrollMoneyValue_(row[PAYROLL_EMPLOYEE_COLUMNS.vehicleAllowance]);
      const housingDeduction = parsePayrollMoneyValue_(row[PAYROLL_EMPLOYEE_COLUMNS.housingDeduction]);
      const municipalTax = parsePayrollMoneyValue_(row[PAYROLL_EMPLOYEE_COLUMNS.municipalTax]);
      const withholding = normalizePayrollWithholdingType_(row[PAYROLL_EMPLOYEE_COLUMNS.withholding]);
      const withholdingLabel = formatPayrollWithholdingLabel_(withholding);
      const dependentCount = normalizePayrollDependentCount_(row[PAYROLL_EMPLOYEE_COLUMNS.dependentCount]);
      const withholdingCategory = normalizePayrollWithholdingCategory_(row[PAYROLL_EMPLOYEE_COLUMNS.withholdingCategory]);
      const withholdingCategoryLabel = formatPayrollWithholdingCategoryLabel_(withholdingCategory);
      const employmentPeriodType = normalizePayrollWithholdingPeriodType_(row[PAYROLL_EMPLOYEE_COLUMNS.withholdingPeriodType]);
      const employmentPeriodLabel = formatPayrollWithholdingPeriodLabel_(employmentPeriodType);
      const transportationType = normalizePayrollTransportationType_(row[PAYROLL_EMPLOYEE_COLUMNS.transportationType]);
      const transportationLabel = formatPayrollTransportationLabel_(transportationType);
      const transportationAmount = parsePayrollMoneyValue_(row[PAYROLL_EMPLOYEE_COLUMNS.transportationAmount]);
      const commissionLogic = normalizePayrollCommissionLogicType_(row[PAYROLL_EMPLOYEE_COLUMNS.commissionLogic]);
      const commissionLabel = formatPayrollCommissionLabel_(commissionLogic);
      const note = String(row[PAYROLL_EMPLOYEE_COLUMNS.note] || '').trim();
      const updatedAt = parseDateValue_(row[PAYROLL_EMPLOYEE_COLUMNS.updatedAt]);
      const record = {
        rowIndex,
        id,
        name,
        email,
        normalizedEmail,
        base,
        baseKey,
        employmentType,
        employmentLabel,
        baseSalary,
        hourlyWage,
        personalAllowance,
        grade,
        qualificationAllowance,
        vehicleAllowance,
        housingDeduction,
        municipalTax,
        withholding,
        withholdingLabel,
        withholdingCategory,
        withholdingCategoryLabel,
        withholdingPeriodType: employmentPeriodType,
        withholdingPeriodLabel: employmentPeriodLabel,
        employmentPeriodType,
        employmentPeriodLabel,
        dependentCount,
        transportationType,
        transportationLabel,
        transportationAmount,
        commissionLogic,
        commissionLabel,
        note,
        updatedAt
      };
      records.push(record);
      if (id) {
        mapById.set(id, record);
      }
      if (normalizedEmail) {
        mapByEmail.set(normalizedEmail, record);
      }
    });
  }
  return { sheet, records, mapById, mapByEmail };
}

function readPayrollRoleAssignments_(){
  const sheet = ensurePayrollRoleSheet_();
  const lastRow = sheet.getLastRow();
  const width = PAYROLL_ROLE_SHEET_HEADER.length;
  const mapByEmail = new Map();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    values.forEach(row => {
      const email = normalizeEmailKey_(row[0]);
      if (!email) return;
      const role = normalizePayrollRoleKey_(row[1]);
      if (!role) return;
      const base = String(row[2] || '').trim();
      const baseKey = normalizePayrollBaseKey_(base);
      mapByEmail.set(email, { email, role, base, baseKey });
    });
  }
  return { sheet, mapByEmail };
}

function readPayrollPayoutEvents_(){
  const sheet = ensurePayrollPayoutEventSheet_();
  const lastRow = sheet.getLastRow();
  const width = PAYROLL_PAYOUT_EVENT_HEADER.length;
  const records = [];
  const mapById = new Map();
  const mapByEmployeeId = new Map();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    values.forEach((row, idx) => {
      const isEmpty = row.every(cell => cell === '' || cell == null);
      if (isEmpty) return;
      const rowIndex = idx + 2;
      let id = String(row[PAYROLL_PAYOUT_EVENT_COLUMNS.id] || '').trim();
      if (!id) {
        id = Utilities.getUuid();
        sheet.getRange(rowIndex, PAYROLL_PAYOUT_EVENT_COLUMN_INDEX.id).setValue(id);
      }
      const employeeId = String(row[PAYROLL_PAYOUT_EVENT_COLUMNS.employeeId] || '').trim();
      if (!employeeId) return;
      const payoutType = normalizePayrollPayoutType_(row[PAYROLL_PAYOUT_EVENT_COLUMNS.payoutType]);
      const fiscalYearCandidate = Number(row[PAYROLL_PAYOUT_EVENT_COLUMNS.fiscalYear]);
      const fiscalYear = Number.isFinite(fiscalYearCandidate) ? fiscalYearCandidate : null;
      const monthKey = normalizePayrollMonthKey_(row[PAYROLL_PAYOUT_EVENT_COLUMNS.monthKey]);
      const periodStart = parseDateValue_(row[PAYROLL_PAYOUT_EVENT_COLUMNS.periodStart]);
      const periodEnd = parseDateValue_(row[PAYROLL_PAYOUT_EVENT_COLUMNS.periodEnd]);
      const payDate = parseDateValue_(row[PAYROLL_PAYOUT_EVENT_COLUMNS.payDate]);
      const title = String(row[PAYROLL_PAYOUT_EVENT_COLUMNS.title] || '').trim();
      const status = String(row[PAYROLL_PAYOUT_EVENT_COLUMNS.status] || '').trim() || 'draft';
      const details = parseJsonColumnValue_(row[PAYROLL_PAYOUT_EVENT_COLUMNS.detailsJson]);
      const insurance = parseJsonColumnValue_(row[PAYROLL_PAYOUT_EVENT_COLUMNS.insuranceJson]);
      const adjustments = parseJsonColumnValue_(row[PAYROLL_PAYOUT_EVENT_COLUMNS.adjustmentJson]);
      const metadata = parseJsonColumnValue_(row[PAYROLL_PAYOUT_EVENT_COLUMNS.metadataJson]);
      const updatedAt = parseDateValue_(row[PAYROLL_PAYOUT_EVENT_COLUMNS.updatedAt]);
      const record = {
        id,
        employeeId,
        payoutType,
        fiscalYear,
        monthKey: monthKey || '',
        periodStart,
        periodEnd,
        payDate,
        title,
        status,
        details,
        insurance,
        adjustments,
        metadata,
        updatedAt,
        rowIndex
      };
      records.push(record);
      mapById.set(id, record);
      if (!mapByEmployeeId.has(employeeId)) {
        mapByEmployeeId.set(employeeId, []);
      }
      mapByEmployeeId.get(employeeId).push(record);
    });
  }
  return { sheet, records, mapById, mapByEmployeeId };
}

function readPayrollAnnualSummaries_(){
  const sheet = ensurePayrollAnnualSummarySheet_();
  const lastRow = sheet.getLastRow();
  const width = PAYROLL_ANNUAL_SUMMARY_HEADER.length;
  const records = [];
  const mapById = new Map();
  const mapByEmployeeYear = new Map();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    values.forEach((row, idx) => {
      const isEmpty = row.every(cell => cell === '' || cell == null);
      if (isEmpty) return;
      const rowIndex = idx + 2;
      let id = String(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.id] || '').trim();
      if (!id) {
        id = Utilities.getUuid();
        sheet.getRange(rowIndex, PAYROLL_ANNUAL_SUMMARY_COLUMN_INDEX.id).setValue(id);
      }
      const employeeId = String(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.employeeId] || '').trim();
      if (!employeeId) return;
      const fiscalYearCandidate = Number(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.fiscalYear]);
      const fiscalYear = Number.isFinite(fiscalYearCandidate) ? fiscalYearCandidate : null;
      const taxableAmount = parsePayrollMoneyValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.taxableAmount]);
      const nonTaxableAmount = parsePayrollMoneyValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.nonTaxableAmount]);
      const socialInsurance = parsePayrollMoneyValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.socialInsurance]);
      const employmentInsurance = parsePayrollMoneyValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.employmentInsurance]);
      const withholdingTax = parsePayrollMoneyValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.withholdingTax]);
      const municipalTax = parsePayrollMoneyValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.municipalTax]);
      const yearEndAdjustment = parsePayrollMoneyValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.yearEndAdjustment]);
      const bonusAmount = parsePayrollMoneyValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.bonusAmount]);
      const payoutCountCandidate = Number(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.payoutCount]);
      const payoutCount = Number.isFinite(payoutCountCandidate) ? payoutCountCandidate : null;
      const summary = parseJsonColumnValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.summaryJson]);
      const metadata = parseJsonColumnValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.metadataJson]);
      const updatedAt = parseDateValue_(row[PAYROLL_ANNUAL_SUMMARY_COLUMNS.updatedAt]);
      const record = {
        id,
        employeeId,
        fiscalYear,
        taxableAmount,
        nonTaxableAmount,
        socialInsurance,
        employmentInsurance,
        withholdingTax,
        municipalTax,
        yearEndAdjustment,
        bonusAmount,
        payoutCount,
        summary,
        metadata,
        updatedAt,
        rowIndex
      };
      records.push(record);
      mapById.set(id, record);
      const key = employeeId + '::' + (fiscalYear != null ? String(fiscalYear) : '');
      mapByEmployeeYear.set(key, record);
    });
  }
  return { sheet, records, mapById, mapByEmployeeYear };
}

function resolvePayrollUserAccess_(){
  const email = (Session.getActiveUser() || {}).getEmail() || '';
  const normalizedEmail = normalizeEmailKey_(email);
  const roles = readPayrollRoleAssignments_();
  let entry = normalizedEmail ? roles.mapByEmail.get(normalizedEmail) : null;
  let role = entry && entry.role ? entry.role : '';
  let base = entry && entry.base ? entry.base : '';
  let baseKey = entry && entry.baseKey ? entry.baseKey : '';
  if (!role) {
    if (isAdminUser_()) {
      role = 'representative';
    } else {
      role = 'none';
    }
  }
  return {
    email: normalizedEmail || '',
    role,
    base,
    baseKey,
    isRepresentative: role === 'representative',
    isManager: role === 'manager'
  };
}

function requirePayrollAccess_(){
  const access = resolvePayrollUserAccess_();
  if (access.role !== 'representative' && access.role !== 'manager') {
    throw new Error('給与マスタの権限がありません。');
  }
  if (access.role === 'manager' && !access.baseKey) {
    throw new Error('管理者の拠点が未設定です。');
  }
  return access;
}

function filterPayrollEmployeesByAccess_(records, access){
  if (!Array.isArray(records)) return [];
  if (!access || access.role === 'representative') {
    return records.slice();
  }
  const baseKey = normalizePayrollBaseKey_(access.baseKey || access.base || '');
  if (!baseKey) return [];
  return records.filter(record => normalizePayrollBaseKey_(record && (record.baseKey || record.base)) === baseKey);
}

function assertPayrollEmployeeAccessible_(employee, access){
  if (!employee) {
    throw new Error('従業員が見つかりません。');
  }
  if (!access || access.role === 'representative') {
    return;
  }
  const employeeBaseKey = normalizePayrollBaseKey_(employee.baseKey || employee.base || '');
  const accessBaseKey = normalizePayrollBaseKey_(access.baseKey || access.base || '');
  if (!accessBaseKey) {
    throw new Error('管理者の拠点が未設定です。');
  }
  if (!employeeBaseKey || employeeBaseKey !== accessBaseKey) {
    throw new Error('この従業員のデータへアクセスする権限がありません。');
  }
}

function assertPayrollMonthIsEditable_(monthKey){
  const normalized = normalizePayrollMonthKey_(monthKey);
  if (!normalized) return;
  const targetDate = createDateFromKey_(normalized + '-01');
  if (!targetDate) return;
  const now = new Date();
  const previousMonthStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  if (targetDate.getTime() < previousMonthStart.getTime()) {
    throw new Error('前月以前の給与明細は修正できません。');
  }
}

function readPayrollSocialInsuranceStandards_(){
  const sheet = ensurePayrollSocialInsuranceStandardSheet_();
  const lastRow = sheet.getLastRow();
  const width = PAYROLL_SOCIAL_INSURANCE_STANDARD_HEADER.length;
  const records = [];
  const mapById = new Map();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    values.forEach((row, idx) => {
      const isEmpty = row.every(cell => cell === '' || cell == null);
      if (isEmpty) return;
      const rowIndex = idx + 2;
      let id = String(row[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.id] || '').trim();
      if (!id) {
        id = Utilities.getUuid();
        sheet.getRange(rowIndex, PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMN_INDEX.id).setValue(id);
      }
      const grade = String(row[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.grade] || '').trim();
      const monthlyAmount = parsePayrollMoneyValue_(row[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.monthlyAmount]);
      const lowerBound = parsePayrollMoneyValue_(row[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.lowerBound]);
      const upperBound = parsePayrollMoneyValue_(row[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.upperBound]);
      const note = String(row[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.note] || '').trim();
      const updatedAt = parseDateValue_(row[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.updatedAt]);
      const record = { id, grade, monthlyAmount, lowerBound, upperBound, note, updatedAt, rowIndex };
      records.push(record);
      mapById.set(id, record);
    });
  }
  records.sort((a, b) => {
    const amountA = Number(a && a.monthlyAmount) || 0;
    const amountB = Number(b && b.monthlyAmount) || 0;
    if (amountA === amountB) {
      const gradeA = String(a && a.grade || '');
      const gradeB = String(b && b.grade || '');
      return gradeA.localeCompare(gradeB, 'ja');
    }
    return amountA - amountB;
  });
  return { sheet, records, mapById };
}

function readPayrollSocialInsuranceOverrides_(){
  const sheet = ensurePayrollSocialInsuranceOverrideSheet_();
  const lastRow = sheet.getLastRow();
  const width = PAYROLL_SOCIAL_INSURANCE_OVERRIDE_HEADER.length;
  const records = [];
  const mapById = new Map();
  const mapByEmployeeMonth = new Map();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    values.forEach((row, idx) => {
      const isEmpty = row.every(cell => cell === '' || cell == null);
      if (isEmpty) return;
      const rowIndex = idx + 2;
      let id = String(row[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.id] || '').trim();
      if (!id) {
        id = Utilities.getUuid();
        sheet.getRange(rowIndex, PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMN_INDEX.id).setValue(id);
      }
      const employeeId = String(row[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.employeeId] || '').trim();
      const monthKey = normalizePayrollMonthKey_(row[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.monthKey]);
      const grade = String(row[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.grade] || '').trim();
      const monthlyAmount = parsePayrollMoneyValue_(row[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.monthlyAmount]);
      const note = String(row[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.note] || '').trim();
      const updatedAt = parseDateValue_(row[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.updatedAt]);
      if (!employeeId || !monthKey) return;
      const record = { id, employeeId, monthKey, grade, monthlyAmount, note, updatedAt, rowIndex };
      records.push(record);
      mapById.set(id, record);
      const key = employeeId + '::' + monthKey;
      mapByEmployeeMonth.set(key, record);
    });
  }
  return { sheet, records, mapById, mapByEmployeeMonth };
}

function buildPayrollSocialInsuranceStandardResponse_(record){
  if (!record) return null;
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  return {
    id: record.id,
    grade: record.grade,
    monthlyAmount: record.monthlyAmount,
    lowerBound: record.lowerBound,
    upperBound: record.upperBound,
    note: record.note,
    updatedAt: record.updatedAt ? formatIsoStringWithOffset_(record.updatedAt, tz) : null
  };
}

function buildPayrollSocialInsuranceOverrideResponse_(record){
  if (!record) return null;
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  return {
    id: record.id,
    employeeId: record.employeeId,
    monthKey: record.monthKey,
    grade: record.grade,
    monthlyAmount: record.monthlyAmount,
    note: record.note,
    updatedAt: record.updatedAt ? formatIsoStringWithOffset_(record.updatedAt, tz) : null
  };
}

function sanitizePayrollRateValue_(value){
  if (value == null || value === '') return null;
  const num = Number(value);
  if (!Number.isFinite(num) || num < 0) return null;
  return Math.round(num * 1000000) / 1000000;
}

function getPayrollSocialInsuranceRates_(){
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PAYROLL_SOCIAL_INSURANCE_RATE_PROPERTY_KEY);
  if (!raw) {
    return { ...PAYROLL_SOCIAL_INSURANCE_RATE_DEFAULTS };
  }
  try {
    const parsed = JSON.parse(raw);
    const merged = { ...PAYROLL_SOCIAL_INSURANCE_RATE_DEFAULTS };
    Object.keys(PAYROLL_SOCIAL_INSURANCE_RATE_DEFAULTS).forEach(key => {
      const sanitized = sanitizePayrollRateValue_(parsed[key]);
      if (sanitized != null) {
        merged[key] = sanitized;
      }
    });
    return merged;
  } catch (err) {
    Logger.log('[getPayrollSocialInsuranceRates_] Failed to parse overrides: ' + err);
    return { ...PAYROLL_SOCIAL_INSURANCE_RATE_DEFAULTS };
  }
}

function savePayrollSocialInsuranceRates_(payload){
  const merged = { ...PAYROLL_SOCIAL_INSURANCE_RATE_DEFAULTS };
  Object.keys(PAYROLL_SOCIAL_INSURANCE_RATE_DEFAULTS).forEach(key => {
    const sanitized = sanitizePayrollRateValue_(payload && payload[key]);
    if (sanitized != null) {
      merged[key] = sanitized;
    }
  });
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PAYROLL_SOCIAL_INSURANCE_RATE_PROPERTY_KEY, JSON.stringify(merged));
  return merged;
}

function sanitizePayrollTaxTableUrl_(url){
  const text = String(url || '').trim();
  if (!text) return '';
  if (!/^https?:\/\//i.test(text)) return '';
  return text;
}

function normalizePayrollIncomeTaxRows_(rows){
  if (!Array.isArray(rows)) return [];
  return rows
    .map(row => (Array.isArray(row) ? row.map(cell => stripBom_(String(cell || '')).trim()) : []))
    .filter(row => Array.isArray(row));
}

function stripBom_(text){
  if (typeof text !== 'string') return text;
  return text.replace(/^\ufeff/, '');
}

function combinePayrollIncomeTaxRows_(koRows, otsuRows){
  const normalizedKo = normalizePayrollIncomeTaxRows_(koRows);
  const normalizedOtsu = normalizePayrollIncomeTaxRows_(otsuRows);
  const combined = normalizedKo.slice();
  if (normalizedOtsu.length) {
    if (combined.length) combined.push([]);
    combined.push(...normalizedOtsu);
  }
  return combined;
}

function buildPayrollIncomeTaxTablesFromSheet_(rows){
  const ko = {};
  const otsu = [];
  const header = Array.isArray(rows && rows[0]) ? rows[0] : [];
  const defaultDependents = Array.from({ length: 8 }).map((_, idx) => ({ count: idx, index: idx + 2 }));
  let dependentsColumns = header
    .map((cell, idx) => ({ count: parsePayrollIncomeTaxDependentHeader_(cell), index: idx }))
    .filter(entry => Number.isFinite(entry.count) && entry.count <= 7);
  if (!dependentsColumns.length) {
    dependentsColumns = defaultDependents;
  }
  const seenDependents = new Set();
  dependentsColumns = dependentsColumns.filter(entry => {
    if (seenDependents.has(entry.count)) return false;
    seenDependents.add(entry.count);
    return true;
  });
  const otsuHeaderIndex = header.findIndex(cell => String(cell || '').includes('乙'));
  const otsuIndex = otsuHeaderIndex >= 0
    ? otsuHeaderIndex
    : Math.max(...dependentsColumns.map(entry => entry.index)) + 1;
  rows.slice(1).forEach(row => {
    if (!row) return;
    const lowerCandidate = parsePayrollMoneyValue_(row[0]);
    const upperCandidate = parsePayrollMoneyValue_(row[1]);
    if (!Number.isFinite(lowerCandidate)) return;
    const min = Math.max(0, lowerCandidate);
    const upperExclusive = Number.isFinite(upperCandidate) ? Math.max(min, upperCandidate) : Infinity;
    const max = upperExclusive === Infinity ? Infinity : Math.max(min, upperExclusive - 1);
    dependentsColumns.forEach(col => {
      const tax = parsePayrollMoneyValue_(row[col.index]);
      if (!Number.isFinite(tax)) return;
      const key = String(col.count);
      if (!ko[key]) ko[key] = [];
      ko[key].push({ min, max, tax });
    });
    const otsuTax = parsePayrollMoneyValue_(row[otsuIndex]);
    if (Number.isFinite(otsuTax)) {
      otsu.push({ min, max, tax: otsuTax });
    }
  });
  if (!Object.keys(ko).length && !otsu.length) {
    throw new Error('所得税税額表シートのデータを読み込めませんでした。');
  }
  return { ko, otsu };
}

function splitPayrollIncomeTaxRows_(rows){
  const normalized = normalizePayrollIncomeTaxRows_(rows);
  const parsedKo = parsePayrollKoIncomeTaxTable_(normalized);
  const endIndex = parsedKo && parsedKo.endIndex ? parsedKo.endIndex : 0;
  const koRows = normalized
    .slice(0, endIndex || normalized.length)
    .filter(row => row.some(cell => String(cell || '').trim() !== ''));
  const otsuRows = normalized
    .slice(endIndex || normalized.length)
    .filter(row => row.some(cell => String(cell || '').trim() !== ''));
  if (!koRows.length && !otsuRows.length && normalized.length) {
    return { koRows: normalized, otsuRows: [] };
  }
  return { koRows, otsuRows };
}

function parsePayrollIncomeTaxRange_(label, previousUpper){
  const numbers = String(label || '').match(/[0-9,]+/g);
  if (!numbers || !numbers.length) return null;
  const parsed = numbers.map(val => parsePayrollMoneyValue_(val)).filter(num => Number.isFinite(num));
  if (!parsed.length) return null;
  const upper = parsed.length >= 2 ? parsed[parsed.length - 1] : parsed[0];
  const lower = parsed.length >= 2
    ? parsed[0]
    : (Number.isFinite(previousUpper) ? previousUpper + 1 : 0);
  if (!Number.isFinite(upper) || !Number.isFinite(lower) || upper < 0) return null;
  const min = Math.max(0, lower);
  if (upper < min) return null;
  return { min, max: upper };
}

function parsePayrollIncomeTaxDependentHeader_(cell){
  const match = String(cell || '').match(/(\d+)/);
  if (!match) return null;
  const num = Number(match[1]);
  return Number.isFinite(num) ? num : null;
}

function parsePayrollKoIncomeTaxTable_(rows){
  const ko = {};
  let endIndex = 0;
  for (let i = 0; i < rows.length; i++) {
    const headerRow = rows[i] || [];
    const dependents = headerRow.map((cell, idx) => ({
      count: parsePayrollIncomeTaxDependentHeader_(cell),
      index: idx
    })).filter(item => Number.isFinite(item.count) && item.count <= 10);
    if (!dependents.length || dependents.every(item => item.index === 0)) continue;
    let previousUpper = null;
    for (let r = i + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const isBlank = row.every(cell => String(cell || '').trim() === '');
      if (isBlank) {
        endIndex = r;
        break;
      }
      const range = parsePayrollIncomeTaxRange_(row[0], previousUpper);
      if (!range) continue;
      previousUpper = range.max;
      dependents.forEach(dep => {
        const tax = parsePayrollMoneyValue_(row[dep.index]);
        if (!Number.isFinite(tax)) return;
        const key = String(dep.count);
        if (!ko[key]) ko[key] = [];
        ko[key].push({ min: range.min, max: range.max, tax });
      });
    }
    break;
  }
  if (!endIndex && Object.keys(ko).length) {
    endIndex = rows.length;
  }
  return { ko, endIndex };
}

function parsePayrollOtsuIncomeTaxTable_(rows, startIndex){
  const otsu = [];
  let started = false;
  let previousUpper = null;
  for (let i = startIndex; i < rows.length; i++) {
    const row = rows[i] || [];
    const first = String(row[0] || '').trim();
    const second = String(row[1] || '').trim();
    const isHeaderRow = (!started && (first.includes('金額') || first.includes('乙') || second.includes('税額')));
    if (!started) {
      if (isHeaderRow) {
        started = true;
        continue;
      }
      if (!first || !second) continue;
      const taxCandidate = parsePayrollMoneyValue_(second);
      if (!Number.isFinite(taxCandidate)) continue;
      started = true;
    }
    if (started) {
      if (!first && !second) {
        if (otsu.length) break;
        continue;
      }
      const range = parsePayrollIncomeTaxRange_(first, previousUpper);
      const tax = parsePayrollMoneyValue_(second);
      if (!range || !Number.isFinite(tax)) continue;
      previousUpper = range.max;
      otsu.push({ min: range.min, max: range.max, tax });
    }
  }
  return otsu;
}

function parsePayrollExplicitIncomeTaxRows_(rows){
  const ko = {};
  const otsu = [];
  rows.forEach(row => {
    if (!Array.isArray(row) || row.length < 5) return;
    const categoryRaw = String(row[0] || '').trim().toLowerCase();
    const dependents = normalizePayrollDependentCount_(row[1]);
    const min = parsePayrollMoneyValue_(row[2]);
    const max = parsePayrollMoneyValue_(row[3]);
    const tax = parsePayrollMoneyValue_(row[4]);
    if (!Number.isFinite(min) || !Number.isFinite(max) || !Number.isFinite(tax)) return;
    if (categoryRaw === 'otsu' || categoryRaw === '乙' || categoryRaw === '乙欄') {
      otsu.push({ min, max, tax });
      return;
    }
    const key = String(dependents);
    if (!ko[key]) ko[key] = [];
    ko[key].push({ min, max, tax });
  });
  return { ko, otsu };
}

function buildPayrollIncomeTaxTablesFromCsv_(rawRows){
  const rows = normalizePayrollIncomeTaxRows_(rawRows);
  const explicit = parsePayrollExplicitIncomeTaxRows_(rows);
  const parsedKo = parsePayrollKoIncomeTaxTable_(rows);
  const otsuStartIndex = Math.max(parsedKo.endIndex || 0, 0);
  const parsedOtsu = parsePayrollOtsuIncomeTaxTable_(rows, otsuStartIndex);
  const mergedKo = { ...explicit.ko };
  Object.keys(parsedKo.ko || {}).forEach(key => {
    if (!mergedKo[key]) mergedKo[key] = [];
    mergedKo[key] = mergedKo[key].concat(parsedKo.ko[key]);
  });
  const mergedOtsu = (explicit.otsu || []).concat(parsedOtsu || []);
  Object.keys(mergedKo).forEach(key => {
    mergedKo[key] = mergedKo[key]
      .filter(entry => Number.isFinite(entry.min) && Number.isFinite(entry.max) && Number.isFinite(entry.tax))
      .sort((a, b) => a.min - b.min);
  });
  const normalizedOtsu = mergedOtsu
    .filter(entry => Number.isFinite(entry.min) && Number.isFinite(entry.max) && Number.isFinite(entry.tax))
    .sort((a, b) => a.min - b.min);
  return { ko: mergedKo, otsu: normalizedOtsu };
}

function persistPayrollIncomeTaxRawSections_(sections, metadata){
  const koRows = sections && Array.isArray(sections.koRows) ? sections.koRows : [];
  const otsuRows = sections && Array.isArray(sections.otsuRows) ? sections.otsuRows : [];
  const props = PropertiesService.getScriptProperties();
  props.setProperty(WITHHOLDING_TAX_TABLE_KO_PROPERTY_KEY, JSON.stringify(koRows));
  props.setProperty(WITHHOLDING_TAX_TABLE_OTSU_PROPERTY_KEY, JSON.stringify(otsuRows));
  props.setProperty(WITHHOLDING_TAX_TABLE_META_PROPERTY_KEY, JSON.stringify({
    fetchedAt: (metadata && metadata.fetchedAt) || new Date().toISOString(),
    fileName: metadata && metadata.fileName ? metadata.fileName : '',
    koDependents: koRows.length,
    otsuRows: otsuRows.length
  }));
  return { koRows, otsuRows };
}

function getPayrollIncomeTaxTables_(options){
  const forceRefresh = options && options.forceRefresh;
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    const cached = cache.get(PAYROLL_INCOME_TAX_CACHE_KEY);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (err) {
        Logger.log('[getPayrollIncomeTaxTables_] Failed to parse cache: ' + (err && err.message ? err.message : err));
      }
    }
  }

  const sheet = ss().getSheetByName(PAYROLL_INCOME_TAX_SHEET_NAME);
  if (!sheet) {
    throw new Error('所得税税額表シートが存在しません。管理者に確認してください。');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('所得税税額表シートに税額データがありません。');
  }
  const lastCol = Math.max(10, sheet.getLastColumn());
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const tables = buildPayrollIncomeTaxTablesFromSheet_(values);
  const payload = {
    ...tables,
    fetchedAt: new Date().toISOString(),
    fileName: PAYROLL_INCOME_TAX_SHEET_NAME
  };
  try {
    cache.put(PAYROLL_INCOME_TAX_CACHE_KEY, JSON.stringify(payload), PAYROLL_INCOME_TAX_CACHE_TTL_SECONDS);
  } catch (err) {
    Logger.log('[getPayrollIncomeTaxTables_] Failed to cache tables: ' + (err && err.message ? err.message : err));
  }
  return payload;
}

function savePayrollIncomeTaxTables_(tables, metadata, rawSections){
  const payload = {
    ko: tables && tables.ko ? tables.ko : {},
    otsu: tables && Array.isArray(tables.otsu) ? tables.otsu : [],
    fetchedAt: (metadata && metadata.fetchedAt) || new Date().toISOString()
  };
  const stats = {
    koDependents: Object.keys(payload.ko || {}).length,
    otsuRows: Array.isArray(payload.otsu) ? payload.otsu.length : 0
  };
  const meta = {
    fetchedAt: payload.fetchedAt,
    fileName: metadata && metadata.fileName ? metadata.fileName : '',
    koDependents: stats.koDependents,
    otsuRows: stats.otsuRows
  };
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PAYROLL_INCOME_TAX_TABLE_CACHE_PROPERTY_KEY, JSON.stringify(payload));
  if (rawSections && rawSections.koRows) {
    props.setProperty(WITHHOLDING_TAX_TABLE_KO_PROPERTY_KEY, JSON.stringify(rawSections.koRows));
  }
  if (rawSections && rawSections.otsuRows) {
    props.setProperty(WITHHOLDING_TAX_TABLE_OTSU_PROPERTY_KEY, JSON.stringify(rawSections.otsuRows));
  }
  props.setProperty(WITHHOLDING_TAX_TABLE_META_PROPERTY_KEY, JSON.stringify(meta));
  return { ...payload, stats, fileName: meta.fileName };
}

function lookupPayrollIncomeTax_(tables, category, taxableAmount, dependents){
  if (!tables || !Number.isFinite(taxableAmount) || taxableAmount <= 0) return null;
  const amount = Math.max(0, Math.floor(taxableAmount));
  if (category === 'otsu') {
    const match = (tables.otsu || []).find(entry => amount >= entry.min && amount <= entry.max);
    return match ? Math.max(0, Math.round(match.tax)) : null;
  }
  const key = String(normalizePayrollDependentCount_(dependents));
  const candidates = (tables.ko && (tables.ko[key] || tables.ko['fuyou' + key])) || [];
  const match = candidates.find(entry => amount >= entry.min && amount <= entry.max);
  if (match) return Math.max(0, Math.round(match.tax));
  return null;
}

function getPayrollIncomeTaxSettings_(){
  const cached = getPayrollIncomeTaxTables_();
  const koKeys = cached && cached.ko ? Object.keys(cached.ko) : [];
  return {
    fetchedAt: cached && cached.fetchedAt ? cached.fetchedAt : null,
    koDependents: koKeys.length,
    otsuRows: Array.isArray(cached && cached.otsu) ? cached.otsu.length : 0,
    fileName: cached && cached.fileName ? cached.fileName : ''
  };
}

function estimatePayrollMonthlyCompensation_(employee){
  if (!employee) return 0;
  if (Number.isFinite(employee.baseSalary) && employee.baseSalary > 0) {
    return Math.round(employee.baseSalary);
  }
  const hourly = Number(employee.hourlyWage) || 0;
  const monthlyFromHourly = hourly > 0 ? hourly * 160 : 0;
  const allowances = ['personalAllowance','qualificationAllowance','vehicleAllowance']
    .map(key => Number(employee[key]) || 0)
    .reduce((sum, val) => sum + Math.max(0, val), 0);
  return Math.round(monthlyFromHourly + allowances);
}

function estimatePayrollTaxableCompensation_(employee, options){
  if (!employee) return 0;
  const allowances = ['personalAllowance','qualificationAllowance','vehicleAllowance']
    .map(key => Math.max(0, Number(employee[key]) || 0))
    .reduce((sum, val) => sum + val, 0);
  const gradeAllowance = Math.max(0, Number(options && options.gradeAmount) || 0);
  const extra = allowances + gradeAllowance;
  const baseSalary = Number(employee.baseSalary);
  if (Number.isFinite(baseSalary) && baseSalary > 0) {
    return Math.round(baseSalary + extra);
  }
  const hourlyWage = Number(employee.hourlyWage) || 0;
  const estimatedHourly = hourlyWage > 0 ? hourlyWage * 160 : 0;
  return Math.round(estimatedHourly + extra);
}

function buildPayrollTaxableBreakdown_(employee, options){
  const gradeAmount = Math.max(0, Number(options && options.gradeAmount) || 0);
  const transportationType = normalizePayrollTransportationType_(employee && employee.transportationType);
  const transportationAmount = Math.max(0, Number(employee && employee.transportationAmount) || 0);
  const nonTaxableTransportation = transportationType === 'none' ? 0 : transportationAmount;
  const housingDeduction = Math.max(0, Number(employee && employee.housingDeduction) || 0);
  const grossCompensation = Math.max(0, estimatePayrollTaxableCompensation_(employee, { gradeAmount }) + transportationAmount);
  const taxableCompensation = Math.max(0, grossCompensation - nonTaxableTransportation - housingDeduction);
  return { grossCompensation, taxableCompensation, nonTaxableTransportation, housingDeduction };
}

function calculatePayrollIncomeTaxFromAnnual_(annualTaxable){
  const taxable = Math.max(0, Number(annualTaxable) || 0);
  const brackets = [
    { upper: 1950000, rate: 0.05, deduction: 0 },
    { upper: 3300000, rate: 0.1, deduction: 97500 },
    { upper: 6950000, rate: 0.2, deduction: 427500 },
    { upper: 9000000, rate: 0.23, deduction: 636000 },
    { upper: 18000000, rate: 0.33, deduction: 1536000 },
    { upper: 40000000, rate: 0.4, deduction: 2796000 },
    { upper: Infinity, rate: 0.45, deduction: 4796000 }
  ];
  const bracket = brackets.find(entry => taxable <= entry.upper) || brackets[brackets.length - 1];
  const base = taxable * bracket.rate - bracket.deduction;
  const incomeTax = Math.max(0, base);
  const surtax = incomeTax * 0.021; // 復興特別所得税
  return Math.floor(incomeTax + surtax);
}

function calculatePayrollKoWithholdingTax_(monthlyTaxable, dependents){
  const clampedDependents = normalizePayrollDependentCount_(dependents);
  const annualTaxable = Math.max(0, (Number(monthlyTaxable) || 0) * 12 - 480000 - (clampedDependents * 380000));
  if (annualTaxable <= 0) return 0;
  const annualTax = calculatePayrollIncomeTaxFromAnnual_(annualTaxable);
  return Math.floor(annualTax / 12);
}

function calculatePayrollWithholdingTax_(employee, options){
  if (!employee || employee.withholding !== 'required') return 0;
  const employmentType = employee.employmentType || '';
  const breakdown = buildPayrollTaxableBreakdown_(employee, options);
  const periodDays = Number(options && options.payPeriodDays);
  const dependents = normalizePayrollDependentCount_(employee.dependentCount);
  const category = normalizePayrollWithholdingCategory_(employee.withholdingCategory);
  const employmentPeriodType = normalizePayrollWithholdingPeriodType_(employee.employmentPeriodType || employee.withholdingPeriodType);
  const socialInsuranceEmployeeTotal = Math.max(0, Number(options && options.socialInsuranceEmployeeTotal) || 0);
  const baseBeforeInsurance = options && options.taxableCompensation != null
    ? Math.max(0, Number(options.taxableCompensation))
    : breakdown.taxableCompensation;
  const taxableBase = Math.max(0, baseBeforeInsurance - socialInsuranceEmployeeTotal);
  if (!Number.isFinite(taxableBase) || taxableBase <= 0) return 0;

  if (employmentType === 'contractor') {
    const rateCandidate = Number(options && options.withholdingRate);
    const rate = Number.isFinite(rateCandidate) && rateCandidate > 0
      ? rateCandidate
      : PAYROLL_WITHHOLDING_TAX_RATE;
    return Math.floor(taxableBase * rate);
  }

  const payDays = Number.isFinite(periodDays) && periodDays > 0 ? periodDays : 30;
  const normalizedBase = employmentPeriodType === 'daily'
    ? (taxableBase / payDays) * 30
    : taxableBase;
  const taxTables = getPayrollIncomeTaxTables_();
  const tableAmount = lookupPayrollIncomeTax_(taxTables, category, normalizedBase, dependents);
  if (tableAmount != null) {
    return tableAmount;
  }
  throw new Error('税額表に該当する行がありません。');
}

function buildPayrollDeductionEntry_(employee, options){
  if (!employee) return null;
  const gradeAmount = Number(options && options.gradeAmount) || 0;
  const payPeriodDays = Number(options && options.payPeriodDays);
  const breakdown = buildPayrollTaxableBreakdown_(employee, { gradeAmount });
  const socialInsuranceEmployeeTotal = Math.max(0, Number(options && options.socialInsuranceEmployeeTotal) || 0);
  const taxableCompensation = Math.max(0, breakdown.taxableCompensation - socialInsuranceEmployeeTotal);
  const withholdingAmount = calculatePayrollWithholdingTax_(employee, {
    taxableCompensation,
    withholdingRate: PAYROLL_WITHHOLDING_TAX_RATE,
    gradeAmount,
    payPeriodDays,
    socialInsuranceEmployeeTotal
  });
  const housingDeduction = breakdown.housingDeduction;
  const municipalTax = Math.max(0, Number(employee.municipalTax) || 0);
  const components = [];
  const suffix = employee.employmentType === 'contractor'
    ? '（業務委託）'
    : `（${formatPayrollWithholdingCategoryLabel_(employee.withholdingCategory)}${employee.dependentCount != null ? `・扶養${normalizePayrollDependentCount_(employee.dependentCount)}` : ''}）`;
  const rate = employee.employmentType === 'contractor' ? PAYROLL_WITHHOLDING_TAX_RATE : null;
  components.push({ type: 'withholding', label: '所得税' + suffix, amount: withholdingAmount, rate });
  if (housingDeduction > 0) {
    components.push({ type: 'housing', label: '社宅費控除', amount: housingDeduction });
  }
  if (municipalTax > 0) {
    components.push({ type: 'municipal', label: '住民税', amount: municipalTax });
  }
  const totalAmount = withholdingAmount + housingDeduction + municipalTax;
  return {
    employeeId: employee.id || '',
    employeeName: employee.name || '',
    employmentType: employee.employmentType || '',
    employmentLabel: employee.employmentLabel || '',
    taxableCompensation,
    withholdingAmount,
    housingDeduction,
    municipalTax,
    totalAmount,
    components,
    gradeAllowanceAmount: gradeAmount
  };
}

function matchPayrollSocialInsuranceStandard_(amount, standards){
  if (!Array.isArray(standards) || standards.length === 0) return null;
  const target = Number(amount);
  const sorted = standards.slice().sort((a, b) => {
    const aVal = Number(a && a.lowerBound != null ? a.lowerBound : a && a.monthlyAmount) || 0;
    const bVal = Number(b && b.lowerBound != null ? b.lowerBound : b && b.monthlyAmount) || 0;
    if (aVal === bVal) {
      return (Number(a && a.monthlyAmount) || 0) - (Number(b && b.monthlyAmount) || 0);
    }
    return aVal - bVal;
  });
  if (!Number.isFinite(target) || target <= 0) {
    return sorted[0];
  }
  let fallback = sorted[sorted.length - 1];
  for (let i = 0; i < sorted.length; i++) {
    const entry = sorted[i];
    const lower = Number(entry && entry.lowerBound);
    const upper = Number(entry && entry.upperBound);
    const min = Number.isFinite(lower) ? lower : 0;
    const max = Number.isFinite(upper) ? upper : Infinity;
    if (target >= min && target <= max) {
      return entry;
    }
    if (target > max) {
      fallback = entry;
    }
  }
  return fallback || sorted[sorted.length - 1];
}

function calculatePayrollSocialInsuranceContribution_(standardAmount, rates, options){
  const amount = Number(standardAmount) || 0;
  const rateSet = rates || PAYROLL_SOCIAL_INSURANCE_RATE_DEFAULTS;
  const compensationBaseCandidate = options && options.compensationAmount != null ? Number(options.compensationAmount) : amount;
  const compensationBase = Number.isFinite(compensationBaseCandidate) && compensationBaseCandidate > 0
    ? compensationBaseCandidate
    : 0;
  const round = (val) => Math.round((Number(val) || 0));
  const healthEmployee = round(amount * (rateSet.healthEmployee || 0));
  const healthEmployer = round(amount * (rateSet.healthEmployer || 0));
  const pensionEmployee = round(amount * (rateSet.pensionEmployee || 0));
  const pensionEmployer = round(amount * (rateSet.pensionEmployer || 0));
  const nursingEmployee = round(amount * (rateSet.nursingEmployee || 0));
  const nursingEmployer = round(amount * (rateSet.nursingEmployer || 0));
  const childEmployee = round(amount * (rateSet.childEmployee || 0));
  const childEmployer = round(amount * (rateSet.childEmployer || 0));
  const employmentEmployee = round(compensationBase * (rateSet.employmentEmployee || 0));
  const employmentEmployer = round(compensationBase * (rateSet.employmentEmployer || 0));
  const employeeTotal = healthEmployee + pensionEmployee + nursingEmployee + childEmployee + employmentEmployee;
  const employerTotal = healthEmployer + pensionEmployer + nursingEmployer + childEmployer + employmentEmployer;
  return {
    healthEmployee,
    healthEmployer,
    pensionEmployee,
    pensionEmployer,
    nursingEmployee,
    nursingEmployer,
    childEmployee,
    childEmployer,
    employmentEmployee,
    employmentEmployer,
    employeeTotal,
    employerTotal
  };
}

function buildPayrollSocialInsuranceSummaryEntry_(employee, options){
  const monthKey = options && options.monthKey;
  const rates = options && options.rates;
  const standard = options && options.standard;
  const override = options && options.override;
  const monthlyCompensation = estimatePayrollMonthlyCompensation_(employee);
  const applied = override && override.monthlyAmount != null ? override : standard;
  const appliedAmount = applied && applied.monthlyAmount != null ? applied.monthlyAmount : (standard && standard.monthlyAmount != null ? standard.monthlyAmount : 0);
  const contributions = calculatePayrollSocialInsuranceContribution_(appliedAmount, rates, {
    compensationAmount: monthlyCompensation
  });
  return {
    employeeId: employee && employee.id ? employee.id : '',
    employeeName: employee && employee.name ? employee.name : '',
    employmentType: employee && employee.employmentType ? employee.employmentType : '',
    monthKey,
    compensationAmount: monthlyCompensation,
    matchedGrade: standard && standard.grade ? standard.grade : '',
    matchedStandardAmount: standard && standard.monthlyAmount != null ? standard.monthlyAmount : null,
    appliedGrade: applied && applied.grade ? applied.grade : (standard && standard.grade ? standard.grade : ''),
    appliedAmount,
    overrideId: override && override.id ? override.id : '',
    overrideNote: override && override.note ? override.note : '',
    isOverride: Boolean(override && override.id),
    contributions
  };
}

function buildPayrollGradeResponse_(record){
  if (!record) return null;
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  return {
    id: record.id,
    name: record.name,
    amount: record.amount,
    note: record.note,
    updatedAt: record.updatedAt ? formatIsoStringWithOffset_(record.updatedAt, tz) : null
  };
}

function buildPayrollEmployeeResponse_(record, gradeMatch){
  if (!record) return null;
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  const normalizedDependents = normalizePayrollDependentCount_(record.dependentCount);
  const normalizedCategory = normalizePayrollWithholdingCategory_(record.withholdingCategory);
  const normalizedEmploymentPeriod = normalizePayrollWithholdingPeriodType_(record.employmentPeriodType || record.withholdingPeriodType);
  const employmentPeriodLabel = formatPayrollWithholdingPeriodLabel_(normalizedEmploymentPeriod);
  return {
    id: record.id,
    name: record.name,
    email: record.email,
    base: record.base,
    employmentType: record.employmentType,
    employmentLabel: record.employmentLabel,
    baseSalary: record.baseSalary,
    hourlyWage: record.hourlyWage,
    personalAllowance: record.personalAllowance,
    grade: record.grade,
    gradeId: gradeMatch ? gradeMatch.id : null,
    gradeAmount: gradeMatch && gradeMatch.amount != null ? gradeMatch.amount : null,
    gradeMasterName: gradeMatch ? gradeMatch.name : null,
    gradeMasterNote: gradeMatch ? gradeMatch.note : null,
    qualificationAllowance: record.qualificationAllowance,
    vehicleAllowance: record.vehicleAllowance,
    housingDeduction: record.housingDeduction,
    municipalTax: record.municipalTax,
    withholding: record.withholding,
    withholdingLabel: record.withholdingLabel,
    withholdingCategory: normalizedCategory,
    withholdingCategoryLabel: record.withholdingCategoryLabel,
    withholdingPeriodType: normalizedEmploymentPeriod,
    withholdingPeriodLabel: employmentPeriodLabel,
    employmentPeriodType: normalizedEmploymentPeriod,
    employmentPeriodLabel,
    dependentCount: normalizedDependents,
    transportationType: record.transportationType,
    transportationLabel: record.transportationLabel,
    transportationAmount: record.transportationAmount,
    commissionLogic: record.commissionLogic,
    commissionLabel: record.commissionLabel,
    note: record.note,
    updatedAt: record.updatedAt ? formatIsoStringWithOffset_(record.updatedAt, tz) : null
  };
}

function buildPayrollPayoutEventResponse_(record, employee){
  if (!record) return null;
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  const employeeName = employee && employee.name
    ? employee.name
    : (record.metadata && record.metadata.employeeName ? record.metadata.employeeName : '');
  return {
    id: record.id,
    employeeId: record.employeeId,
    employeeName,
    payoutType: record.payoutType,
    payoutTypeLabel: formatPayrollPayoutLabel_(record.payoutType),
    fiscalYear: record.fiscalYear,
    monthKey: record.monthKey || '',
    monthLabel: formatPayrollMonthLabel_(record.monthKey || '', tz),
    periodStart: record.periodStart ? Utilities.formatDate(record.periodStart, tz, 'yyyy-MM-dd') : null,
    periodEnd: record.periodEnd ? Utilities.formatDate(record.periodEnd, tz, 'yyyy-MM-dd') : null,
    payDate: record.payDate ? Utilities.formatDate(record.payDate, tz, 'yyyy-MM-dd') : null,
    title: record.title || '',
    status: record.status || 'draft',
    details: record.details || null,
    insurance: record.insurance || null,
    adjustments: record.adjustments || null,
    metadata: record.metadata || null,
    updatedAt: record.updatedAt ? formatIsoStringWithOffset_(record.updatedAt, tz) : null
  };
}

function buildPayrollAnnualSummaryResponse_(record, employee){
  if (!record) return null;
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  const employeeName = employee && employee.name
    ? employee.name
    : (record.metadata && record.metadata.employeeName ? record.metadata.employeeName : '');
  return {
    id: record.id,
    employeeId: record.employeeId,
    employeeName,
    fiscalYear: record.fiscalYear,
    taxableAmount: record.taxableAmount,
    nonTaxableAmount: record.nonTaxableAmount,
    socialInsurance: record.socialInsurance,
    employmentInsurance: record.employmentInsurance,
    withholdingTax: record.withholdingTax,
    municipalTax: record.municipalTax,
    yearEndAdjustment: record.yearEndAdjustment,
    bonusAmount: record.bonusAmount,
    payoutCount: record.payoutCount,
    summary: record.summary || null,
    metadata: record.metadata || null,
    updatedAt: record.updatedAt ? formatIsoStringWithOffset_(record.updatedAt, tz) : null
  };
}

function payrollGetIncomeTaxSettings(){
  return wrapPayrollResponse_('payrollGetIncomeTaxSettings', () => {
    requirePayrollAccess_();
    const settings = getPayrollIncomeTaxSettings_();
    return { ok: true, settings };
  });
}

function payrollSaveIncomeTaxCsvUrl(payload){
  return wrapPayrollResponse_('payrollSaveIncomeTaxCsvUrl', () => {
    return { ok: false, reason: 'deprecated', message: '税額表のアップロードに切り替わりました。CSVファイルをアップロードしてください。' };
  });
}


function payrollReloadIncomeTaxTables(payload){
  return wrapPayrollResponse_('payrollReloadIncomeTaxTables', () => {
    requirePayrollAccess_();
    const tables = getPayrollIncomeTaxTables_({ forceRefresh: true });
    const stats = {
      koDependents: Object.keys(tables && tables.ko ? tables.ko : {}).length,
      otsuRows: Array.isArray(tables && tables.otsu) ? tables.otsu.length : 0
    };
    return {
      ok: true,
      fetchedAt: tables && tables.fetchedAt ? tables.fetchedAt : null,
      stats,
      fileName: tables && tables.fileName ? tables.fileName : '',
      message: '所得税税額表を再読み込みしました。'
    };
  });
}


function payrollUploadIncomeTaxCsv(file){
  return wrapPayrollResponse_('payrollUploadIncomeTaxCsv', () => {
    requirePayrollAccess_();
    return { ok: false, reason: 'validation', message: '所得税税額表シートを直接更新してください。CSVのアップロードは不要です。' };
  });
}

function payrollListEmployees(){
  return wrapPayrollResponse_('payrollListEmployees', () => {
    const access = requirePayrollAccess_();
    const context = readPayrollEmployeeRecords_();
    const gradeContext = readPayrollGradeRecords_();
    const scoped = filterPayrollEmployeesByAccess_(context.records, access);
    const list = scoped.slice().sort((a, b) => {
      const nameA = (a && a.name) ? a.name.toString() : '';
      const nameB = (b && b.name) ? b.name.toString() : '';
      return nameA.localeCompare(nameB, 'ja');
    }).map(record => {
      const normalized = normalizePayrollGradeName_(record && record.grade);
      const gradeMatch = normalized ? gradeContext.mapByNormalizedName.get(normalized) : null;
      return buildPayrollEmployeeResponse_(record, gradeMatch);
    });
    return { ok: true, employees: list };
  });
}

function payrollListDeductionSummary(){
  return wrapPayrollResponse_('payrollListDeductionSummary', () => {
    const access = requirePayrollAccess_();
    const employeeContext = readPayrollEmployeeRecords_();
    const gradeContext = readPayrollGradeRecords_();
    const insuranceStandards = readPayrollSocialInsuranceStandards_();
    const insuranceOverrides = readPayrollSocialInsuranceOverrides_();
    const insuranceRates = getPayrollSocialInsuranceRates_();
    const monthKey = normalizePayrollMonthKey_(new Date());
    const scoped = filterPayrollEmployeesByAccess_(employeeContext.records, access);
    const entries = scoped.map(record => {
      const normalized = normalizePayrollGradeName_(record && record.grade);
      const gradeMatch = normalized ? gradeContext.mapByNormalizedName.get(normalized) : null;
      const gradeAmount = gradeMatch && gradeMatch.amount != null ? gradeMatch.amount : 0;
      const overrideKey = record.id + '::' + monthKey;
      const override = insuranceOverrides.mapByEmployeeMonth.get(overrideKey);
      const standard = matchPayrollSocialInsuranceStandard_(estimatePayrollMonthlyCompensation_(record), insuranceStandards.records);
      const insuranceEntry = buildPayrollSocialInsuranceSummaryEntry_(record, { monthKey, rates: insuranceRates, standard, override });
      const socialTotal = insuranceEntry && insuranceEntry.contributions ? Number(insuranceEntry.contributions.employeeTotal) || 0 : 0;
      return buildPayrollDeductionEntry_(record, { gradeAmount, socialInsuranceEmployeeTotal: socialTotal, payPeriodDays: 30 });
    }).filter(entry => entry && entry.totalAmount > 0);
    entries.sort((a, b) => (a.employeeName || '').localeCompare(b.employeeName || '', 'ja'));
    const totals = entries.reduce((acc, entry) => {
      acc.withholding += entry.withholdingAmount || 0;
      acc.housing += entry.housingDeduction || 0;
      acc.municipal += entry.municipalTax || 0;
      acc.total += entry.totalAmount || 0;
      return acc;
    }, { withholding: 0, housing: 0, municipal: 0, total: 0 });
    const tz = getConfig('timezone') || 'Asia/Tokyo';
    return {
      ok: true,
      generatedAt: formatIsoStringWithOffset_(new Date(), tz),
      totals,
      employees: entries
    };
  });
}

function payrollListGrades(){
  return wrapPayrollResponse_('payrollListGrades', () => {
    const context = readPayrollGradeRecords_();
    const list = context.records.slice().sort((a, b) => {
      const nameA = (a && a.name) ? a.name.toString() : '';
      const nameB = (b && b.name) ? b.name.toString() : '';
      return nameA.localeCompare(nameB, 'ja');
    }).map(record => buildPayrollGradeResponse_(record));
    return { ok: true, grades: list };
  });
}

function payrollSaveEmployee(payload){
  return wrapPayrollResponse_('payrollSaveEmployee', () => {
    const access = requirePayrollAccess_();
    const name = String(payload && payload.name || '').trim();
    if (!name) {
      return { ok: false, reason: 'validation', message: '氏名を入力してください。' };
    }
    const email = String(payload && payload.email || '').trim();
    const employmentType = normalizePayrollEmploymentType_(payload && payload.employmentType);
    const dependentCountInput = payload && payload.dependentCount;
    const dependentCount = normalizePayrollDependentCount_(dependentCountInput);
    if (!employmentType) {
      return { ok: false, reason: 'validation', message: '雇用区分を選択してください。' };
    }
    if (dependentCount < 0 || dependentCount > 7) {
      return { ok: false, reason: 'validation', message: '扶養人数は0〜7の範囲で入力してください。' };
    }
    let base = String(payload && payload.base || '').trim();
    let baseKey = normalizePayrollBaseKey_(base);
    const baseSalary = parsePayrollMoneyValue_(payload && payload.baseSalary);
    const hourlyWage = parsePayrollMoneyValue_(payload && payload.hourlyWage);
    const personalAllowance = parsePayrollMoneyValue_(payload && payload.personalAllowance);
    const qualificationAllowance = parsePayrollMoneyValue_(payload && payload.qualificationAllowance);
    const vehicleAllowance = parsePayrollMoneyValue_(payload && payload.vehicleAllowance);
    const housingDeduction = parsePayrollMoneyValue_(payload && payload.housingDeduction);
    const municipalTax = parsePayrollMoneyValue_(payload && payload.municipalTax);
    const transportationType = normalizePayrollTransportationType_(payload && payload.transportationType);
    const transportationAmount = parsePayrollMoneyValue_(payload && payload.transportationAmount);
    const commissionLogic = normalizePayrollCommissionLogicType_(payload && payload.commissionLogic);
    const withholding = normalizePayrollWithholdingType_(payload && payload.withholding);
    const withholdingCategoryInput = payload && payload.withholdingCategory;
    const withholdingCategory = normalizePayrollWithholdingCategory_(withholdingCategoryInput);
    const employmentPeriodType = normalizePayrollWithholdingPeriodType_(payload && payload.employmentPeriodType);
    const grade = String(payload && payload.grade || '').trim();
    const note = String(payload && payload.note || '').trim();
    let id = String(payload && payload.id || '').trim();

    const context = readPayrollEmployeeRecords_();
    const sheet = context.sheet;
    let rowIndex;
    let existing = null;
    if (id) {
      existing = context.mapById.get(id);
      if (!existing) {
        return { ok: false, reason: 'not_found', message: '従業員が見つかりません。' };
      }
      rowIndex = existing.rowIndex;
    } else {
      id = Utilities.getUuid();
      rowIndex = sheet.getLastRow() + 1;
    }

    if (access.role === 'manager') {
      const managerBaseKey = normalizePayrollBaseKey_(access.baseKey || access.base || '');
      if (!managerBaseKey) {
        return { ok: false, reason: 'forbidden', message: '管理者の拠点が未設定です。' };
      }
      if (existing) {
        const existingBaseKey = normalizePayrollBaseKey_(existing.base);
        if (existingBaseKey && existingBaseKey !== managerBaseKey) {
          return { ok: false, reason: 'forbidden', message: '自拠点以外の従業員は編集できません。' };
        }
      }
      if (baseKey && baseKey !== managerBaseKey) {
        return { ok: false, reason: 'forbidden', message: '自拠点以外の従業員は編集できません。' };
      }
      if (!base) {
        base = access.base || '';
      }
      baseKey = managerBaseKey;
    }

    if (withholding === 'required') {
      if (dependentCountInput === '' || dependentCountInput == null) {
        return { ok: false, reason: 'validation', message: '扶養人数を入力してください。' };
      }
      if (!String(withholdingCategoryInput || '').trim()) {
        return { ok: false, reason: 'validation', message: '源泉所得税の甲乙区分を選択してください。' };
      }
    }

    const values = new Array(PAYROLL_EMPLOYEE_SHEET_HEADER.length).fill('');
    values[PAYROLL_EMPLOYEE_COLUMNS.id] = id;
    values[PAYROLL_EMPLOYEE_COLUMNS.name] = name;
    values[PAYROLL_EMPLOYEE_COLUMNS.email] = email;
    values[PAYROLL_EMPLOYEE_COLUMNS.base] = base;
    values[PAYROLL_EMPLOYEE_COLUMNS.employmentType] = formatPayrollEmploymentLabel_(employmentType);
    values[PAYROLL_EMPLOYEE_COLUMNS.baseSalary] = baseSalary != null ? baseSalary : '';
    values[PAYROLL_EMPLOYEE_COLUMNS.hourlyWage] = hourlyWage != null ? hourlyWage : '';
    values[PAYROLL_EMPLOYEE_COLUMNS.personalAllowance] = personalAllowance != null ? personalAllowance : '';
    values[PAYROLL_EMPLOYEE_COLUMNS.grade] = grade;
    values[PAYROLL_EMPLOYEE_COLUMNS.qualificationAllowance] = qualificationAllowance != null ? qualificationAllowance : '';
    values[PAYROLL_EMPLOYEE_COLUMNS.vehicleAllowance] = vehicleAllowance != null ? vehicleAllowance : '';
    values[PAYROLL_EMPLOYEE_COLUMNS.housingDeduction] = housingDeduction != null ? housingDeduction : '';
    values[PAYROLL_EMPLOYEE_COLUMNS.municipalTax] = municipalTax != null ? municipalTax : '';
    values[PAYROLL_EMPLOYEE_COLUMNS.withholding] = formatPayrollWithholdingLabel_(withholding);
    values[PAYROLL_EMPLOYEE_COLUMNS.transportationType] = formatPayrollTransportationLabel_(transportationType);
    values[PAYROLL_EMPLOYEE_COLUMNS.transportationAmount] = transportationAmount != null ? transportationAmount : '';
    values[PAYROLL_EMPLOYEE_COLUMNS.commissionLogic] = formatPayrollCommissionLabel_(commissionLogic);
    values[PAYROLL_EMPLOYEE_COLUMNS.note] = note;
    values[PAYROLL_EMPLOYEE_COLUMNS.updatedAt] = new Date();
    values[PAYROLL_EMPLOYEE_COLUMNS.dependentCount] = dependentCount;
    values[PAYROLL_EMPLOYEE_COLUMNS.withholdingCategory] = formatPayrollWithholdingCategoryLabel_(withholdingCategory);
    values[PAYROLL_EMPLOYEE_COLUMNS.withholdingPeriodType] = formatPayrollWithholdingPeriodLabel_(employmentPeriodType);

    sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);

    const refreshed = readPayrollEmployeeRecords_();
    const saved = refreshed.mapById.get(id);
    const gradeContext = readPayrollGradeRecords_();
    const gradeMatch = saved ? gradeContext.mapByNormalizedName.get(normalizePayrollGradeName_(saved.grade)) : null;
    return { ok: true, employee: buildPayrollEmployeeResponse_(saved, gradeMatch) };
  });
}

function payrollDeleteEmployee(payload){
  return wrapPayrollResponse_('payrollDeleteEmployee', () => {
    const id = String(payload && payload.id || '').trim();
    if (!id) {
      return { ok: false, reason: 'validation', message: '削除する従業員を指定してください。' };
    }
    const access = requirePayrollAccess_();
    const context = readPayrollEmployeeRecords_();
    const record = context.mapById.get(id);
    if (!record) {
      return { ok: false, reason: 'not_found', message: '対象の従業員が見つかりません。' };
    }
    assertPayrollEmployeeAccessible_(record, access);
    context.sheet.deleteRow(record.rowIndex);
    return { ok: true };
  });
}

function payrollSaveGrade(payload){
  return wrapPayrollResponse_('payrollSaveGrade', () => {
    const name = String(payload && payload.name || '').trim();
    if (!name) {
      return { ok: false, reason: 'validation', message: '役職/等級名を入力してください。' };
    }
    const amount = parsePayrollMoneyValue_(payload && payload.amount);
    const note = String(payload && payload.note || '').trim();
    let id = String(payload && payload.id || '').trim();

    const context = readPayrollGradeRecords_();
    const normalized = normalizePayrollGradeName_(name);
    if (normalized) {
      const duplicate = context.mapByNormalizedName.get(normalized);
      if (duplicate && (!id || duplicate.id !== id)) {
        return { ok: false, reason: 'validation', message: '同じ名前の役職/等級が既に存在します。' };
      }
    }

    let rowIndex;
    const sheet = context.sheet;
    if (id) {
      const existing = context.mapById.get(id);
      if (!existing) {
        return { ok: false, reason: 'not_found', message: '対象の役職/等級が見つかりません。' };
      }
      rowIndex = existing.rowIndex;
    } else {
      id = Utilities.getUuid();
      rowIndex = sheet.getLastRow() + 1;
    }

    const values = new Array(PAYROLL_GRADE_SHEET_HEADER.length).fill('');
    values[PAYROLL_GRADE_COLUMNS.id] = id;
    values[PAYROLL_GRADE_COLUMNS.name] = name;
    values[PAYROLL_GRADE_COLUMNS.amount] = amount != null ? amount : '';
    values[PAYROLL_GRADE_COLUMNS.note] = note;
    values[PAYROLL_GRADE_COLUMNS.updatedAt] = new Date();

    sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);

    const refreshed = readPayrollGradeRecords_();
    const saved = refreshed.mapById.get(id);
    return { ok: true, grade: buildPayrollGradeResponse_(saved) };
  });
}

function payrollDeleteGrade(payload){
  return wrapPayrollResponse_('payrollDeleteGrade', () => {
    const id = String(payload && payload.id || '').trim();
    if (!id) {
      return { ok: false, reason: 'validation', message: '削除する役職/等級を指定してください。' };
    }
    const context = readPayrollGradeRecords_();
    const record = context.mapById.get(id);
    if (!record) {
      return { ok: false, reason: 'not_found', message: '対象の役職/等級が見つかりません。' };
    }
    context.sheet.deleteRow(record.rowIndex);
    return { ok: true };
  });
}

function payrollCalculateCommissionSummary(payload){
  return wrapPayrollResponse_('payrollCalculateCommissionSummary', () => {
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    const { year, month, start, end } = resolvePayrollCommissionMonthRange_(payload && payload.month);
    const weekRanges = buildPayrollMonthWeekRanges_(start, end);
    const countsMap = collectTreatmentCountsByStaffForRange_(start, end, tz);
    const access = requirePayrollAccess_();
    const employeeContext = readPayrollEmployeeRecords_();
    const employees = filterPayrollEmployeesByAccess_(employeeContext.records || [], access);
    const summaries = employees.map(record => {
      const emailKey = normalizeEmailKey_(record && record.email);
      const countsEntry = emailKey ? countsMap.get(emailKey) : null;
      return buildPayrollCommissionBreakdown_(record, countsEntry, { tz, weekRanges });
    });
    const label = Utilities.formatDate(start, tz, 'yyyy年MM月');
    return {
      ok: true,
      month: { year, month, label },
      range: {
        startDate: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
        endDateExclusive: Utilities.formatDate(end, tz, 'yyyy-MM-dd')
      },
      weeks: weekRanges.map(range => ({
        startDate: Utilities.formatDate(range.start, tz, 'yyyy-MM-dd'),
        endDate: Utilities.formatDate(new Date(range.end.getTime() - 1), tz, 'yyyy-MM-dd')
      })),
      employees: summaries
    };
  });
}

function buildPayrollAttendanceEntry_(employee, records, staffProfile){
  const normalizedRecords = Array.isArray(records) ? records : [];
  const dailyMap = new Map();
  normalizedRecords.forEach(record => {
    if (!record || !record.date) return;
    if (!dailyMap.has(record.date)) {
      dailyMap.set(record.date, {
        date: record.date,
        workMinutes: 0,
        breakMinutes: 0,
        leaveType: record.metadata && record.metadata.leaveType ? record.metadata.leaveType : '',
        records: []
      });
    }
    const entry = dailyMap.get(record.date);
    entry.workMinutes += Number(record.workMinutes) || 0;
    entry.breakMinutes += Number(record.breakMinutes) || 0;
    entry.records.push(record);
  });

  const workingDays = dailyMap.size;
  const scheduledPerDay = staffProfile && Number.isFinite(staffProfile.defaultShiftMinutes)
    ? staffProfile.defaultShiftMinutes
    : VISIT_ATTENDANCE_DEFAULT_SHIFT_MINUTES;
  let totalWorkMinutes = 0;
  let totalBreakMinutes = 0;
  let overtimeMinutes = 0;
  let autoPaidLeaveMinutes = 0;
  let autoPaidLeaveDays = 0;
  let recordedPaidLeaveDays = 0;

  dailyMap.forEach(entry => {
    const work = Math.max(0, entry.workMinutes || 0);
    const brk = Math.max(0, entry.breakMinutes || 0);
    totalWorkMinutes += work;
    totalBreakMinutes += brk;
    if (entry.leaveType === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE) {
      recordedPaidLeaveDays += 1;
    }
    if (scheduledPerDay > 0) {
      if (work > scheduledPerDay) {
        overtimeMinutes += work - scheduledPerDay;
      } else if (work < scheduledPerDay) {
        autoPaidLeaveMinutes += scheduledPerDay - work;
        autoPaidLeaveDays += 1;
      }
    }
  });

  const scheduledMinutes = workingDays * scheduledPerDay;
  const buildText = (value) => formatMinutesAsTimeText_(Math.max(0, Math.round(value || 0)));

  return {
    employeeId: employee && employee.id ? employee.id : '',
    employeeName: employee && employee.name ? employee.name : '',
    email: employee && employee.email ? employee.email : '',
    employmentType: employee && employee.employmentType ? employee.employmentType : '',
    employmentLabel: employee && employee.employmentLabel ? employee.employmentLabel : '',
    workingDays,
    workMinutes: totalWorkMinutes,
    workText: buildText(totalWorkMinutes),
    workDurationText: formatDurationText_(totalWorkMinutes),
    breakMinutes: totalBreakMinutes,
    breakText: buildText(totalBreakMinutes),
    overtimeMinutes,
    overtimeText: buildText(overtimeMinutes),
    autoPaidLeaveMinutes,
    autoPaidLeaveText: buildText(autoPaidLeaveMinutes),
    autoPaidLeaveDurationText: formatDurationText_(autoPaidLeaveMinutes),
    autoPaidLeaveDays,
    recordedPaidLeaveDays,
    scheduledMinutes,
    scheduledText: buildText(scheduledMinutes),
    scheduledPerDayMinutes: scheduledPerDay,
    hasAttendance: workingDays > 0,
    recordsCount: normalizedRecords.length
  };
}

function payrollListAttendanceSummary(payload){
  return wrapPayrollResponse_('payrollListAttendanceSummary', () => {
    const datasetResult = getUnifiedAttendanceDataset(payload || {});
    if (!datasetResult || datasetResult.ok !== true || !datasetResult.dataset) {
      throw new Error((datasetResult && datasetResult.message) || '勤怠データの取得に失敗しました。');
    }
    const dataset = datasetResult.dataset;
    const staffSettings = readVisitAttendanceStaffSettings_();
    const access = requirePayrollAccess_();
    const employeeContext = readPayrollEmployeeRecords_();
    const employees = filterPayrollEmployeesByAccess_(employeeContext.records || [], access);
    const visitRecordsByEmail = new Map();
    const datasetRecords = Array.isArray(dataset.records) ? dataset.records : [];
    datasetRecords.forEach(record => {
      if (!record || record.staffType !== 'employee') return;
      const normalizedEmail = normalizeEmailKey_(record.staffId);
      if (!normalizedEmail) return;
      if (!visitRecordsByEmail.has(normalizedEmail)) {
        visitRecordsByEmail.set(normalizedEmail, []);
      }
      visitRecordsByEmail.get(normalizedEmail).push(record);
    });

    const employeeEmailSet = new Set();
    employees.forEach(emp => {
      const emailKey = normalizeEmailKey_(emp && emp.email);
      if (emailKey) employeeEmailSet.add(emailKey);
    });

    const employeeSummaries = employees.map(employee => {
      const normalizedEmail = normalizeEmailKey_(employee && employee.email);
      const records = normalizedEmail ? (visitRecordsByEmail.get(normalizedEmail) || []) : [];
      const profile = normalizedEmail ? staffSettings.get(normalizedEmail) : null;
      return buildPayrollAttendanceEntry_(employee, records, profile);
    });

    const unmatchedStaff = [];
    visitRecordsByEmail.forEach((records, email) => {
      if (employeeEmailSet.has(email)) return;
      const workMinutes = records.reduce((sum, record) => sum + (Number(record && record.workMinutes) || 0), 0);
      const breakMinutes = records.reduce((sum, record) => sum + (Number(record && record.breakMinutes) || 0), 0);
      const sample = records[0] || {};
      unmatchedStaff.push({
        email,
        staffName: sample.staffName || email,
        workMinutes,
        workText: formatMinutesAsTimeText_(workMinutes),
        breakMinutes,
        breakText: formatMinutesAsTimeText_(breakMinutes),
        workingDays: records.length,
        durationText: formatDurationText_(workMinutes)
      });
    });

    const tz = dataset.timezone || Session.getScriptTimeZone() || 'Asia/Tokyo';
    const range = dataset.range || {};
    const fromKey = range.from || '';
    const derivedMonthKey = fromKey ? fromKey.slice(0, 7) : '';
    const payloadMonth = payload && (payload.month || payload.monthKey);
    const monthKey = payloadMonth ? String(payloadMonth) : derivedMonthKey;
    const fromDate = fromKey ? createDateFromKey_(fromKey) : null;
    const monthLabel = fromDate ? Utilities.formatDate(fromDate, tz, 'yyyy年M月') : (monthKey || '');

    return {
      ok: true,
      month: {
        key: monthKey,
        label: monthLabel
      },
      range,
      timezone: tz,
      totals: dataset.totals || null,
      systems: dataset.systems || [],
      employees: employeeSummaries,
      unmatchedStaff
    };
  });
}

function payrollListPayoutEvents(payload){
  return wrapPayrollResponse_('payrollListPayoutEvents', () => {
    const access = requirePayrollAccess_();
    const employeeContext = readPayrollEmployeeRecords_();
    const payoutContext = readPayrollPayoutEvents_();
    const employees = filterPayrollEmployeesByAccess_(employeeContext.records || [], access);
    const employeeMap = new Map();
    employees.forEach(employee => {
      if (employee && employee.id) {
        employeeMap.set(employee.id, employee);
      }
    });
    const hasTypeFilter = payload && payload.payoutType != null && String(payload.payoutType).trim() !== '';
    const typeFilter = hasTypeFilter ? normalizePayrollPayoutType_(payload.payoutType) : '';
    const fiscalYearCandidate = Number(payload && (payload.fiscalYear || payload.year));
    const hasFiscalYearFilter = Number.isFinite(fiscalYearCandidate);
    const monthFilter = normalizePayrollMonthKey_(payload && (payload.month || payload.monthKey));
    const hasMonthFilter = Boolean(monthFilter);
    const events = payoutContext.records.filter(record => {
      if (!employeeMap.has(record.employeeId)) return false;
      if (hasTypeFilter && record.payoutType !== typeFilter) return false;
      if (hasFiscalYearFilter && record.fiscalYear !== fiscalYearCandidate) return false;
      if (hasMonthFilter && (record.monthKey || '') !== monthFilter) return false;
      return true;
    }).map(record => {
      const employee = employeeMap.get(record.employeeId) || employeeContext.mapById.get(record.employeeId) || null;
      return buildPayrollPayoutEventResponse_(record, employee);
    });
    return { ok: true, events };
  });
}

function payrollListAnnualSummaries(payload){
  return wrapPayrollResponse_('payrollListAnnualSummaries', () => {
    const access = requirePayrollAccess_();
    const employeeContext = readPayrollEmployeeRecords_();
    const summaryContext = readPayrollAnnualSummaries_();
    const employees = filterPayrollEmployeesByAccess_(employeeContext.records || [], access);
    const employeeMap = new Map();
    employees.forEach(employee => {
      if (employee && employee.id) {
        employeeMap.set(employee.id, employee);
      }
    });
    const fiscalYearCandidate = Number(payload && (payload.fiscalYear || payload.year));
    const hasFiscalYearFilter = Number.isFinite(fiscalYearCandidate);
    const summaries = summaryContext.records.filter(record => {
      if (!employeeMap.has(record.employeeId)) return false;
      if (hasFiscalYearFilter && record.fiscalYear !== fiscalYearCandidate) return false;
      return true;
    }).map(record => {
      const employee = employeeMap.get(record.employeeId) || employeeContext.mapById.get(record.employeeId) || null;
      return buildPayrollAnnualSummaryResponse_(record, employee);
    });
    return { ok: true, summaries };
  });
}

function payrollGetSocialInsuranceSettings(){
  return wrapPayrollResponse_('payrollGetSocialInsuranceSettings', () => {
    const standardsContext = readPayrollSocialInsuranceStandards_();
    const rates = getPayrollSocialInsuranceRates_();
    const standards = standardsContext.records.map(record => buildPayrollSocialInsuranceStandardResponse_(record));
    return { ok: true, standards, rates };
  });
}

function payrollSaveSocialInsuranceStandard(payload){
  return wrapPayrollResponse_('payrollSaveSocialInsuranceStandard', () => {
    const grade = String(payload && payload.grade || '').trim();
    if (!grade) {
      return { ok: false, reason: 'validation', message: '等級を入力してください。' };
    }
    const monthlyAmount = parsePayrollMoneyValue_(payload && payload.monthlyAmount);
    if (monthlyAmount == null) {
      return { ok: false, reason: 'validation', message: '標準報酬月額を入力してください。' };
    }
    const lowerBound = parsePayrollMoneyValue_(payload && payload.lowerBound);
    const upperBound = parsePayrollMoneyValue_(payload && payload.upperBound);
    const note = String(payload && payload.note || '').trim();
    const context = readPayrollSocialInsuranceStandards_();
    const sheet = context.sheet;
    let id = String(payload && payload.id || '').trim();
    let rowIndex;
    if (id) {
      const existing = context.mapById.get(id);
      if (!existing) {
        return { ok: false, reason: 'not_found', message: '対象の等級が見つかりません。' };
      }
      rowIndex = existing.rowIndex;
    } else {
      id = Utilities.getUuid();
      rowIndex = sheet.getLastRow() + 1;
    }
    const values = new Array(PAYROLL_SOCIAL_INSURANCE_STANDARD_HEADER.length).fill('');
    values[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.id] = id;
    values[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.grade] = grade;
    values[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.monthlyAmount] = monthlyAmount;
    values[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.lowerBound] = lowerBound != null ? lowerBound : '';
    values[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.upperBound] = upperBound != null ? upperBound : '';
    values[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.note] = note;
    values[PAYROLL_SOCIAL_INSURANCE_STANDARD_COLUMNS.updatedAt] = new Date();
    sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
    const refreshed = readPayrollSocialInsuranceStandards_();
    const saved = refreshed.mapById.get(id);
    return { ok: true, standard: buildPayrollSocialInsuranceStandardResponse_(saved) };
  });
}

function payrollDeleteSocialInsuranceStandard(payload){
  return wrapPayrollResponse_('payrollDeleteSocialInsuranceStandard', () => {
    const id = String(payload && payload.id || '').trim();
    if (!id) {
      return { ok: false, reason: 'validation', message: '削除対象の等級を指定してください。' };
    }
    const context = readPayrollSocialInsuranceStandards_();
    const record = context.mapById.get(id);
    if (!record) {
      return { ok: false, reason: 'not_found', message: '対象の等級が見つかりません。' };
    }
    context.sheet.deleteRow(record.rowIndex);
    return { ok: true };
  });
}

function payrollSaveSocialInsuranceRates(payload){
  return wrapPayrollResponse_('payrollSaveSocialInsuranceRates', () => {
    const saved = savePayrollSocialInsuranceRates_(payload || {});
    return { ok: true, rates: saved };
  });
}

function payrollListSocialInsuranceSummary(payload){
  return wrapPayrollResponse_('payrollListSocialInsuranceSummary', () => {
    const monthKeyInput = payload && (payload.monthKey || payload.month);
    const normalizedMonthKey = normalizePayrollMonthKey_(monthKeyInput) || normalizePayrollMonthKey_(new Date());
    const tz = getConfig('timezone') || 'Asia/Tokyo';
    const monthDate = normalizedMonthKey ? createDateFromKey_(normalizedMonthKey + '-01') : null;
    const monthLabel = monthDate ? Utilities.formatDate(monthDate, tz, 'yyyy年M月') : normalizedMonthKey;
    const access = requirePayrollAccess_();
    const employeeContext = readPayrollEmployeeRecords_();
    const standardsContext = readPayrollSocialInsuranceStandards_();
    const overridesContext = readPayrollSocialInsuranceOverrides_();
    const rates = getPayrollSocialInsuranceRates_();
    const overrideMap = overridesContext.mapByEmployeeMonth || new Map();
    const employees = filterPayrollEmployeesByAccess_(employeeContext.records || [], access);
    const entries = employees.map(employee => {
      const baseAmount = estimatePayrollMonthlyCompensation_(employee);
      const matchedStandard = matchPayrollSocialInsuranceStandard_(baseAmount, standardsContext.records);
      const overrideKey = employee.id + '::' + normalizedMonthKey;
      const override = overrideMap.get(overrideKey);
      return buildPayrollSocialInsuranceSummaryEntry_(employee, {
        monthKey: normalizedMonthKey,
        rates,
        standard: matchedStandard,
        override
      });
    });
    return {
      ok: true,
      month: { key: normalizedMonthKey, label: monthLabel },
      entries,
      rates
    };
  });
}

function payrollSaveSocialInsuranceOverride(payload){
  return wrapPayrollResponse_('payrollSaveSocialInsuranceOverride', () => {
    const employeeId = String(payload && payload.employeeId || '').trim();
    if (!employeeId) {
      return { ok: false, reason: 'validation', message: '従業員を選択してください。' };
    }
    const access = requirePayrollAccess_();
    const employeeContext = readPayrollEmployeeRecords_();
    const employee = employeeContext.mapById.get(employeeId);
    if (!employee) {
      return { ok: false, reason: 'not_found', message: '従業員が見つかりません。' };
    }
    assertPayrollEmployeeAccessible_(employee, access);
    const monthKey = normalizePayrollMonthKey_(payload && (payload.month || payload.monthKey));
    if (!monthKey) {
      return { ok: false, reason: 'validation', message: '対象月を指定してください。' };
    }
    const monthlyAmount = parsePayrollMoneyValue_(payload && payload.monthlyAmount);
    if (monthlyAmount == null) {
      return { ok: false, reason: 'validation', message: '標準報酬月額を入力してください。' };
    }
    const grade = String(payload && payload.grade || '').trim();
    const note = String(payload && payload.note || '').trim();
    const context = readPayrollSocialInsuranceOverrides_();
    const sheet = context.sheet;
    let id = String(payload && payload.id || '').trim();
    let rowIndex;
    if (id) {
      const existing = context.mapById.get(id);
      if (!existing) {
        return { ok: false, reason: 'not_found', message: '上書きが見つかりません。' };
      }
      rowIndex = existing.rowIndex;
    } else {
      const existingKey = employeeId + '::' + monthKey;
      const existing = context.mapByEmployeeMonth.get(existingKey);
      if (existing) {
        id = existing.id;
        rowIndex = existing.rowIndex;
      } else {
        id = Utilities.getUuid();
        rowIndex = sheet.getLastRow() + 1;
      }
    }
    const values = new Array(PAYROLL_SOCIAL_INSURANCE_OVERRIDE_HEADER.length).fill('');
    values[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.id] = id;
    values[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.employeeId] = employeeId;
    values[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.monthKey] = monthKey;
    values[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.grade] = grade;
    values[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.monthlyAmount] = monthlyAmount;
    values[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.note] = note;
    values[PAYROLL_SOCIAL_INSURANCE_OVERRIDE_COLUMNS.updatedAt] = new Date();
    sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
    const refreshed = readPayrollSocialInsuranceOverrides_();
    const saved = refreshed.mapById.get(id);
    return { ok: true, override: buildPayrollSocialInsuranceOverrideResponse_(saved) };
  });
}

function payrollDeleteSocialInsuranceOverride(payload){
  return wrapPayrollResponse_('payrollDeleteSocialInsuranceOverride', () => {
    const id = String(payload && payload.id || '').trim();
    if (!id) {
      return { ok: false, reason: 'validation', message: '削除対象を選択してください。' };
    }
    const access = requirePayrollAccess_();
    const context = readPayrollSocialInsuranceOverrides_();
    const record = context.mapById.get(id);
    if (!record) {
      return { ok: false, reason: 'not_found', message: '上書きが見つかりません。' };
    }
    const employeeContext = readPayrollEmployeeRecords_();
    const employee = record.employeeId ? employeeContext.mapById.get(record.employeeId) : null;
    if (employee) {
      assertPayrollEmployeeAccessible_(employee, access);
    }
    context.sheet.deleteRow(record.rowIndex);
    return { ok: true };
  });
}

function formatPayrollMonthLabel_(monthKey, tz){
  const normalized = normalizePayrollMonthKey_(monthKey);
  if (!normalized) return '';
  const date = createDateFromKey_(normalized + '-01');
  if (!date) return '';
  const timezone = tz || getConfig('timezone') || 'Asia/Tokyo';
  return Utilities.formatDate(date, timezone, 'yyyy年M月');
}

function buildPayrollPayslipTemplateData_(employee, payload){
  if (!employee) {
    throw new Error('従業員が見つかりません。');
  }
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  const now = new Date();
  const monthKeyInput = payload && (payload.month || payload.monthKey);
  let monthKey = normalizePayrollMonthKey_(monthKeyInput) || normalizePayrollMonthKey_(now);
  let monthDate = monthKey ? createDateFromKey_(monthKey + '-01') : now;
  let defaultStart = monthDate ? new Date(monthDate.getFullYear(), monthDate.getMonth(), 1) : now;
  let defaultEnd = monthDate ? new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0) : now;
  let periodStart = parseDateValue_(payload && (payload.periodStart || payload.periodFrom));
  let periodEnd = parseDateValue_(payload && (payload.periodEnd || payload.periodTo));
  let payday = parseDateValue_(payload && (payload.payday || payload.payDate));
  let issued = parseDateValue_(payload && (payload.issuedAt || payload.issuedDate));
  const payoutEventIdInput = String(payload && payload.payoutEventId || '').trim();
  const payoutTypeInput = payload && (payload.payoutType || payload.type);
  let payoutType = normalizePayrollPayoutType_(payoutTypeInput);
  let payoutEvent = null;
  if (payoutEventIdInput) {
    const payoutContext = readPayrollPayoutEvents_();
    payoutEvent = payoutContext.mapById.get(payoutEventIdInput) || null;
  } else if (payload && payload.payoutEvent) {
    payoutEvent = payload.payoutEvent;
  }
  if (payoutEvent) {
    if (payoutEvent.payoutType) {
      payoutType = normalizePayrollPayoutType_(payoutEvent.payoutType);
    }
    const overrideMonth = normalizePayrollMonthKey_(payoutEvent.monthKey);
    if (overrideMonth) {
      monthKey = overrideMonth;
      monthDate = monthKey ? createDateFromKey_(monthKey + '-01') : now;
      defaultStart = monthDate ? new Date(monthDate.getFullYear(), monthDate.getMonth(), 1) : now;
      defaultEnd = monthDate ? new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0) : now;
    }
    const overridePeriodStart = parseDateValue_(payoutEvent.periodStart || payoutEvent.periodFrom);
    if (overridePeriodStart) {
      periodStart = overridePeriodStart;
    }
    const overridePeriodEnd = parseDateValue_(payoutEvent.periodEnd || payoutEvent.periodTo);
    if (overridePeriodEnd) {
      periodEnd = overridePeriodEnd;
    }
    const overridePayday = parseDateValue_(payoutEvent.payDate || payoutEvent.payday);
    if (overridePayday) {
      payday = overridePayday;
    }
    const overrideIssued = parseDateValue_(payoutEvent.issuedAt || payoutEvent.issuedDate);
    if (overrideIssued) {
      issued = overrideIssued;
    }
  }
  periodStart = periodStart instanceof Date && !isNaN(periodStart.getTime()) ? periodStart : defaultStart;
  periodEnd = periodEnd instanceof Date && !isNaN(periodEnd.getTime()) ? periodEnd : defaultEnd;
  payday = payday instanceof Date && !isNaN(payday.getTime()) ? payday : new Date(defaultEnd.getFullYear(), defaultEnd.getMonth() + 1, 25);
  issued = issued instanceof Date && !isNaN(issued.getTime()) ? issued : now;
  const payPeriodDays = Math.max(1, Math.round((periodEnd.getTime() - periodStart.getTime()) / (1000 * 60 * 60 * 24)) + 1);
  const payPeriodText = Utilities.formatDate(periodStart, tz, 'yyyy年M月d日') + ' 〜 ' + Utilities.formatDate(periodEnd, tz, 'yyyy年M月d日');
  const paydayText = Utilities.formatDate(payday, tz, 'yyyy年M月d日');
  const issuedText = Utilities.formatDate(issued, tz, 'yyyy年M月d日');
  const payoutTypeLabel = formatPayrollPayoutLabel_(payoutType);
  const gradeContext = readPayrollGradeRecords_();
  const normalizedGrade = normalizePayrollGradeName_(employee && employee.grade);
  const gradeRecord = normalizedGrade ? gradeContext.mapByNormalizedName.get(normalizedGrade) : null;
  const gradeAmount = gradeRecord && gradeRecord.amount != null ? gradeRecord.amount : 0;
  const payoutDetails = payoutEvent && payoutEvent.details ? payoutEvent.details : null;
  const payoutDetailsTitle = payoutDetails && payoutDetails.title ? String(payoutDetails.title).trim() : '';
  const payoutTitle = String(payload && payload.payoutTitle || '').trim()
    || (payoutEvent && payoutEvent.title ? payoutEvent.title : '')
    || payoutDetailsTitle
    || ((payoutType === 'bonus' && formatPayrollMonthLabel_(monthKey, tz))
      ? formatPayrollMonthLabel_(monthKey, tz) + ' 賞与'
      : formatPayrollMonthLabel_(monthKey, tz));
  const replacePayItems = payoutDetails && payoutDetails.replaceDefaultPayItems === true;
  const replaceDeductionItems = payoutDetails && payoutDetails.replaceDefaultDeductionItems === true;
  const detailPayItems = Array.isArray(payoutDetails && payoutDetails.earnings) ? payoutDetails.earnings : [];
  const detailDeductionItems = Array.isArray(payoutDetails && payoutDetails.deductions) ? payoutDetails.deductions : [];

  const payItems = [];
  const appendPayItem = (label, amount, note) => {
    const parsed = parsePayrollMoneyValue_(amount);
    if (parsed == null || parsed === 0) return;
    payItems.push({ label: label || '', amount: parsed, note: note ? String(note).trim() : '' });
  };
  if (!replacePayItems) {
    appendPayItem('基本給', employee.baseSalary);
    if ((employee.baseSalary == null || employee.baseSalary === 0) && employee.hourlyWage) {
      const estimated = Math.round((Number(employee.hourlyWage) || 0) * 160);
      if (estimated > 0) {
        appendPayItem('時間給換算 (160h)', estimated);
      }
    }
    appendPayItem(gradeRecord && gradeRecord.name ? gradeRecord.name : '役職手当', gradeAmount);
    appendPayItem('個別加算', employee.personalAllowance);
    appendPayItem('資格手当', employee.qualificationAllowance);
    appendPayItem('車両手当', employee.vehicleAllowance);
    if (employee.transportationType === 'fixed') {
      appendPayItem('交通費（固定）', employee.transportationAmount);
    } else if (employee.transportationType === 'actual') {
      appendPayItem('交通費（実費）', employee.transportationAmount);
    }
  }
  detailPayItems.forEach(item => {
    if (!item) return;
    const label = String(item.label || '').trim() || '支給';
    appendPayItem(label, item.amount, item.note);
  });
  const extraPayItemsSource = []
    .concat(Array.isArray(payload && payload.payItems) ? payload.payItems : [])
    .concat(Array.isArray(payload && payload.extraPayItems) ? payload.extraPayItems : []);
  extraPayItemsSource.forEach(item => {
    if (!item) return;
    const label = String(item.label || '').trim() || '支給';
    appendPayItem(label, item.amount, item.note);
  });

  const deductionItems = [];
  const appendDeductionItem = (label, amount, note) => {
    const parsed = parsePayrollMoneyValue_(amount);
    if (parsed == null || parsed === 0) return;
    deductionItems.push({ label: label || '', amount: parsed, note: note ? String(note).trim() : '' });
  };
  const insuranceOverride = payoutEvent && payoutEvent.insurance ? payoutEvent.insurance : null;
  let socialInsurance = null;
  if (!replaceDeductionItems) {
    if (insuranceOverride && Array.isArray(insuranceOverride.components)) {
      socialInsurance = insuranceOverride;
      insuranceOverride.components.forEach(component => {
        if (!component) return;
        appendDeductionItem(component.label || '社会保険', component.amount, component.note);
      });
    } else if (monthKey) {
      const standardsContext = readPayrollSocialInsuranceStandards_();
      const overridesContext = readPayrollSocialInsuranceOverrides_();
      const rates = getPayrollSocialInsuranceRates_();
      const overrideKey = employee.id + '::' + monthKey;
      const override = overridesContext.mapByEmployeeMonth.get(overrideKey);
      const standard = matchPayrollSocialInsuranceStandard_(estimatePayrollMonthlyCompensation_(employee), standardsContext.records);
      const insuranceEntry = buildPayrollSocialInsuranceSummaryEntry_(employee, { monthKey, rates, standard, override });
      socialInsurance = insuranceEntry;
      if (insuranceEntry && insuranceEntry.contributions) {
        const contrib = insuranceEntry.contributions;
        const healthTotal = (contrib.healthEmployee || 0) + (contrib.nursingEmployee || 0) + (contrib.childEmployee || 0);
        appendDeductionItem('健康保険', healthTotal);
        appendDeductionItem('厚生年金', contrib.pensionEmployee || 0);
        appendDeductionItem('雇用保険', contrib.employmentEmployee || 0);
      }
    }
    const socialInsuranceEmployeeTotal = socialInsurance && socialInsurance.contributions
      ? Number(socialInsurance.contributions.employeeTotal) || 0
      : 0;
    const deductionEntry = buildPayrollDeductionEntry_(employee, { gradeAmount, socialInsuranceEmployeeTotal, payPeriodDays }) || { components: [] };
    if (Array.isArray(deductionEntry.components)) {
      deductionEntry.components.forEach(component => {
        if (!component) return;
        const type = component.type || '';
        let label = component.label || '';
        if (type === 'withholding') label = '所得税';
        if (type === 'housing') label = '社宅控除';
        if (type === 'municipal') label = '住民税';
        appendDeductionItem(label || '控除', component.amount, component.note);
      });
    }
  } else {
    if (insuranceOverride && Array.isArray(insuranceOverride.components)) {
      socialInsurance = insuranceOverride;
      insuranceOverride.components.forEach(component => {
        if (!component) return;
        appendDeductionItem(component.label || '社会保険', component.amount, component.note);
      });
    } else if (insuranceOverride) {
      socialInsurance = insuranceOverride;
    }
  }
  detailDeductionItems.forEach(item => {
    if (!item) return;
    const label = String(item.label || '').trim() || '控除';
    appendDeductionItem(label, item.amount, item.note);
  });
  const extraDeductionSource = []
    .concat(Array.isArray(payload && payload.deductionItems) ? payload.deductionItems : [])
    .concat(Array.isArray(payload && payload.extraDeductionItems) ? payload.extraDeductionItems : []);
  extraDeductionSource.forEach(item => {
    if (!item) return;
    const label = String(item.label || '').trim() || '控除';
    appendDeductionItem(label, item.amount, item.note);
  });

  const filteredPayItems = payItems.filter(item => Number(item.amount) > 0);
  const filteredDeductionItems = deductionItems.filter(item => Number(item.amount) > 0);
  const sumItems = (list) => list.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
  let totalGross = sumItems(filteredPayItems);
  let totalDeduction = sumItems(filteredDeductionItems);
  const payoutSummary = payoutDetails && payoutDetails.summary ? payoutDetails.summary : null;
  if (payoutSummary) {
    const summaryGross = parsePayrollMoneyValue_(payoutSummary.grossAmount);
    if (summaryGross != null) {
      totalGross = summaryGross;
    }
    const summaryDeduction = parsePayrollMoneyValue_(payoutSummary.deductionAmount);
    if (summaryDeduction != null) {
      totalDeduction = summaryDeduction;
    }
  }
  let netPay = totalGross - totalDeduction;
  if (payoutSummary) {
    const summaryNet = parsePayrollMoneyValue_(payoutSummary.netAmount);
    if (summaryNet != null) {
      netPay = summaryNet;
    }
  }
  const formattedPayItems = filteredPayItems.map(item => ({
    label: item.label,
    amount: item.amount,
    amountText: formatPayrollCurrencyYen_(item.amount),
    note: item.note
  }));
  const formattedDeductionItems = filteredDeductionItems.map(item => ({
    label: item.label,
    amount: item.amount,
    amountText: formatPayrollCurrencyYen_(item.amount),
    note: item.note
  }));

  const folderName = formatPayrollEmployeeFolderName_(employee.name);
  const reiwaLabel = formatJapaneseEraMonthLabel_(monthDate || now, tz);
  const fileBase = sanitizeDriveFileName_(`${employee.name || '従業員'}_給与支払明細_${reiwaLabel}`);
  let noteText = String(payload && payload.note || '').trim() || '勤怠・控除内容は労働基準法および就業規則に基づき算出しています。内容に相違がある場合は5営業日以内に総務までご連絡ください。';
  if (payoutDetails && payoutDetails.note) {
    noteText = String(payoutDetails.note).trim();
  }
  let messageHeading = String(payload && payload.messageHeading || '').trim() || 'Message from BellTree';
  if (payoutDetails && payoutDetails.messageHeading) {
    messageHeading = String(payoutDetails.messageHeading).trim();
  }
  let messageBody = String(payload && payload.messageBody || '').trim() || '「あなたの働きが、べるつりーの理念である “関係者に最善を提供する” ことにつながっています。」\n「いつもありがとうございます。」';
  if (payoutDetails && payoutDetails.messageBody) {
    messageBody = String(payoutDetails.messageBody).trim();
  }
  let netPayNote = String(payload && (payload.netPayNote || payload.bankInfo) || employee.note || '').trim() || '振込予定口座：登録口座';
  if (payoutDetails && payoutDetails.netPayNote) {
    netPayNote = String(payoutDetails.netPayNote).trim();
  }
  const brand = {
    initials: String(payload && payload.brandInitials || '').trim() || 'BT',
    title: String(payload && payload.brandTitle || '').trim() || '給与明細',
    tagline: String(payload && payload.brandTagline || '').trim() || 'PAYROLL STATEMENT'
  };
  const footerCompany = String(payload && payload.footerCompany || '').trim() || 'BellTree';
  const footerAddress = String(payload && payload.footerAddress || '').trim() || '〒192-0372 東京都八王子市下柚木3-7-2-401';
  const footerContact = String(payload && payload.footerContact || '').trim() || '代表 042-682-2839 ｜ belltree@belltree1102.com';
  const resolvedPayoutEventId = payoutEvent && payoutEvent.id ? payoutEvent.id : payoutEventIdInput;
  const monthLabel = formatPayrollMonthLabel_(monthKey, tz);

  return {
    employeeId: employee.id,
    employeeName: employee.name,
    employeeDisplayName: (employee.name || '従業員') + ' 様',
    employeeFolderName: folderName,
    monthKey,
    monthLabel,
    payoutType,
    payoutTypeLabel,
    payoutTitle,
    payoutEventId: resolvedPayoutEventId || '',
    reiwaLabel,
    payPeriodText,
    paydayText,
    totalGross,
    totalGrossText: formatPayrollCurrencyYen_(totalGross),
    totalDeduction,
    totalDeductionText: formatPayrollCurrencyYen_(totalDeduction),
    netPay,
    netPayText: formatPayrollCurrencyYen_(netPay),
    netPayNote,
    payItems: formattedPayItems,
    deductionItems: formattedDeductionItems,
    noteHtml: convertPlainTextToSafeHtml_(noteText),
    messageHeading,
    messageBodyHtml: convertPlainTextToSafeHtml_(messageBody),
    issuedText,
    brand,
    footerCompany,
    footerAddress,
    footerContact,
    socialInsurance,
    fileName: fileBase + '.pdf'
  };
}

function createPayrollPayslipPdf_(payload, options){
  const context = options && options.context ? options.context : readPayrollEmployeeRecords_();
  const access = options && options.access ? options.access : null;
  const employeeId = String(payload && payload.employeeId || '').trim();
  let employee = employeeId ? context.mapById.get(employeeId) : null;
  if (!employee) {
    const name = String(payload && payload.employeeName || '').trim();
    if (name) {
      employee = context.records.find(record => record.name === name);
    }
  }
  if (!employee) {
    throw new Error('従業員が見つかりません。');
  }
  if (access) {
    assertPayrollEmployeeAccessible_(employee, access);
  }
  const templateData = buildPayrollPayslipTemplateData_(employee, payload || {});
  const folder = ensurePayrollEmployeeFolder_(employee.name);
  if (payload && payload.overwriteExisting) {
    deletePayrollPayslipFilesByName_(folder, templateData.fileName);
  }
  const template = HtmlService.createTemplateFromFile('payroll_pdf_family');
  template.payrollPdfData = templateData;
  const html = template.evaluate().getContent();
  const blob = Utilities.newBlob(html, 'text/html', 'payroll_payslip.html').getAs(MimeType.PDF);
  blob.setName(templateData.fileName || '給与明細.pdf');
  const file = folder.createFile(blob);
  return { file, data: templateData, folder };
}

function payrollGeneratePayslipPdf(payload){
  return wrapPayrollResponse_('payrollGeneratePayslipPdf', () => {
    const access = requirePayrollAccess_();
    const context = readPayrollEmployeeRecords_();
    const result = createPayrollPayslipPdf_(payload || {}, { context, access });
    return {
      ok: true,
      fileId: result.file.getId(),
      fileName: result.file.getName(),
      fileUrl: result.file.getUrl(),
      folderName: result.folder.getName(),
      data: result.data
    };
  });
}

function payrollBulkGeneratePayslips(payload){
  return wrapPayrollResponse_('payrollBulkGeneratePayslips', () => {
    const access = requirePayrollAccess_();
    const context = readPayrollEmployeeRecords_();
    const tz = getConfig('timezone') || 'Asia/Tokyo';
    const monthInput = payload && (payload.month || payload.monthKey);
    const normalizedMonth = normalizePayrollMonthKey_(monthInput) || normalizePayrollMonthKey_(new Date());
    assertPayrollMonthIsEditable_(normalizedMonth);
    const employees = filterPayrollEmployeesByAccess_(context.records, access);
    if (!employees.length) {
      throw new Error('対象従業員が見つかりません。');
    }
    const successes = [];
    const failures = [];
    employees.forEach(employee => {
      try {
        const result = createPayrollPayslipPdf_(Object.assign({}, payload, {
          employeeId: employee.id,
          month: normalizedMonth,
          overwriteExisting: true
        }), { context, access });
        successes.push({
          employeeId: employee.id,
          employeeName: employee.name,
          fileId: result.file.getId(),
          fileName: result.file.getName(),
          fileUrl: result.file.getUrl()
        });
      } catch (err) {
        failures.push({
          employeeId: employee.id,
          employeeName: employee.name,
          message: err && err.message ? err.message : String(err)
        });
      }
    });
    return {
      ok: true,
      month: { key: normalizedMonth, label: formatPayrollMonthLabel_(normalizedMonth, tz) },
      summary: { total: employees.length, success: successes.length, failed: failures.length },
      results: successes,
      errors: failures
    };
  });
}

function resolveUnifiedAttendanceRange_(payload){
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  let fromKey = normalizeDateKey_(payload && payload.from, tz);
  let toKey = normalizeDateKey_(payload && payload.to, tz);

  if (!fromKey || !toKey) {
    let resolvedMonth = normalizeYearMonthInput_(payload && payload.year, payload && payload.month);
    if (!resolvedMonth) {
      const monthKeyText = String((payload && (payload.monthKey || payload.month)) || '').trim();
      if (/^\d{4}-\d{1,2}$/.test(monthKeyText)) {
        const parts = monthKeyText.split('-');
        resolvedMonth = normalizeYearMonthInput_(Number(parts[0]), Number(parts[1]));
      }
    }
    if (resolvedMonth) {
      const resolved = resolveMonthlyRangeKeys_(resolvedMonth.year, resolvedMonth.month);
      fromKey = resolved.from;
      toKey = resolved.to;
    }
  }

  if (!fromKey || !toKey) {
    const now = new Date();
    const resolved = resolveMonthlyRangeKeys_(now.getFullYear(), now.getMonth() + 1);
    fromKey = resolved.from;
    toKey = resolved.to;
  }

  if (fromKey > toKey) {
    const tmp = fromKey;
    fromKey = toKey;
    toKey = tmp;
  }

  const fromDate = createDateFromKey_(fromKey);
  const toDate = createDateFromKey_(toKey);
  return { tz, fromKey, toKey, fromDate, toDate };
}

function buildUnifiedVisitAttendanceRecord_(record){
  if (!record) return null;
  const workMinutes = Number.isFinite(record.workMinutes) ? record.workMinutes : 0;
  const breakMinutes = Number.isFinite(record.breakMinutes) ? record.breakMinutes : 0;
  return {
    system: 'visit',
    systemLabel: 'VisitAttendance',
    staffType: 'employee',
    staffId: record.email || '',
    staffName: record.staffName || record.email || '',
    date: record.date || '',
    clockIn: record.start || '',
    clockOut: record.end || '',
    breakMinutes,
    breakText: record.break || formatMinutesAsTimeText_(breakMinutes),
    workMinutes,
    workText: record.work || formatMinutesAsTimeText_(workMinutes),
    durationText: formatDurationText_(workMinutes),
    note: record.breakdown || '',
    metadata: {
      weekday: record.weekday || '',
      breakdown: record.breakdown || '',
      flag: record.flag || '',
      sourceLabel: record.sourceLabel || '',
      leaveType: record.leaveType || '',
      isHourlyStaff: !!record.isHourlyStaff,
      isDailyStaff: !!record.isDailyStaff,
      autoAdjustedEnd: !!record.autoAdjustedEnd,
      autoAdjustmentMessage: record.autoAdjustmentMessage || '',
      rowNumber: Number.isFinite(record.rowNumber) ? record.rowNumber : null,
      originalEndMinutes: Number.isFinite(record.originalEndMinutes) ? record.originalEndMinutes : null
    }
  };
}

function buildUnifiedAlbyteAttendanceRecord_(record, options){
  if (!record) return null;
  const staffContext = options && options.staffContext;
  const shiftContext = options && options.shiftContext;
  const staff = staffContext && staffContext.mapById ? staffContext.mapById.get(record.staffId) : null;
  const fallbackStaff = staff || {
    id: record.staffId || '',
    name: record.staffName || '',
    normalizedName: staff && staff.normalizedName ? staff.normalizedName : normalizeAlbyteName_(record.staffName),
    staffType: staff && staff.staffType ? staff.staffType : 'hourly'
  };
  const view = buildAlbyteAttendanceView_(record, { shiftContext, staff: fallbackStaff });
  if (!view) return null;
  const workMinutes = Number.isFinite(view.workMinutes) ? view.workMinutes : 0;
  const breakMinutes = Number.isFinite(view.breakMinutes) ? view.breakMinutes : 0;
  const unifiedStaffType = view.isDailyStaff ? 'daily' : 'partTime';
  return {
    system: 'albyte',
    systemLabel: 'AlbyteAttendance',
    staffType: unifiedStaffType,
    staffId: view.staffId || fallbackStaff.id || '',
    staffName: view.staffName || fallbackStaff.name || '',
    date: view.date || record.date || '',
    clockIn: view.clockIn || '',
    clockOut: view.clockOut || '',
    breakMinutes,
    breakText: view.breakText || formatMinutesAsTimeText_(breakMinutes),
    workMinutes,
    workText: view.workText || formatMinutesAsTimeText_(workMinutes),
    durationText: view.durationText || formatDurationText_(workMinutes),
    note: view.note || '',
    metadata: {
      autoFlag: view.autoFlag || '',
      log: Array.isArray(view.log) ? view.log : [],
      shiftStart: view.shiftStart || '',
      shiftEnd: view.shiftEnd || '',
      shiftNote: view.shiftNote || '',
      rowIndex: Number.isFinite(view.rowIndex) ? view.rowIndex : null,
      recordId: view.id || record.id || '',
      rawStaffId: record.staffId || '',
      rawStaffName: record.staffName || '',
      isDailyStaff: !!view.isDailyStaff,
      overtimeMinutes: view.isDailyStaff ? view.workMinutes : 0,
      overtimeText: view.isDailyStaff ? (view.overtimeText || view.workText || formatMinutesAsTimeText_(workMinutes)) : '',
      baseShiftEnd: view.baseShiftEnd || '',
      sourceLabel: view.sourceLabel || '通常勤務'
    }
  };
}

function getUnifiedAttendanceDataset(payload){
  const tag = 'getUnifiedAttendanceDataset';
  try {
    const range = resolveUnifiedAttendanceRange_(payload);
    const visitRecords = readVisitAttendanceRecords_({
      startDate: range.fromDate,
      endDate: range.toDate,
      tz: range.tz
    });
    const visitUnified = visitRecords.map(buildUnifiedVisitAttendanceRecord_).filter(Boolean);

    const staffContext = readAlbyteStaffRecords_();
    const shiftContext = readAlbyteShiftRecords_();
    const { records: albyteRecords } = readAlbyteAttendanceRecords_({ fromDateKey: range.fromKey, toDateKey: range.toKey });
    const albyteUnified = albyteRecords.map(record => buildUnifiedAlbyteAttendanceRecord_(record, { staffContext, shiftContext })).filter(Boolean);

    const records = visitUnified.concat(albyteUnified);
    records.sort((a, b) => {
      const dateDiff = (a.date || '').localeCompare(b.date || '');
      if (dateDiff !== 0) return dateDiff;
      const typeDiff = (a.staffType || '').localeCompare(b.staffType || '');
      if (typeDiff !== 0) return typeDiff;
      const nameDiff = (a.staffName || '').localeCompare(b.staffName || '');
      if (nameDiff !== 0) return nameDiff;
      const startDiff = (a.clockIn || '').localeCompare(b.clockIn || '');
      if (startDiff !== 0) return startDiff;
      return (a.system || '').localeCompare(b.system || '');
    });

    let totalWork = 0;
    let totalBreak = 0;
    const systemMap = new Map();
    const staffSummaryMap = new Map();

    const ensureSystemEntry = (key, label) => {
      if (!systemMap.has(key)) {
        systemMap.set(key, { system: key, label: label || key, workMinutes: 0, breakMinutes: 0, records: 0 });
      }
      return systemMap.get(key);
    };

    records.forEach(record => {
      const work = Number.isFinite(record.workMinutes) ? record.workMinutes : 0;
      const breakMinutes = Number.isFinite(record.breakMinutes) ? record.breakMinutes : 0;
      totalWork += work;
      totalBreak += breakMinutes;

      const systemEntry = ensureSystemEntry(record.system || 'unknown', record.systemLabel || record.system || '');
      systemEntry.workMinutes += work;
      systemEntry.breakMinutes += breakMinutes;
      systemEntry.records += 1;

      const staffKey = (record.staffType || '') + '::' + (record.staffId || record.staffName || '');
      if (!staffSummaryMap.has(staffKey)) {
        staffSummaryMap.set(staffKey, {
          staffType: record.staffType || '',
          staffId: record.staffId || '',
          staffName: record.staffName || '',
          workMinutes: 0,
          breakMinutes: 0,
          workingDays: 0,
          systems: new Set()
        });
      }
      const summary = staffSummaryMap.get(staffKey);
      summary.workMinutes += work;
      summary.breakMinutes += breakMinutes;
      if ((record.clockIn && record.clockOut) || work > 0) {
        summary.workingDays += 1;
      }
      if (record.system) {
        summary.systems.add(record.system);
      }
    });

    const systemSummaries = Array.from(systemMap.values()).map(entry => ({
      system: entry.system,
      label: entry.label,
      workMinutes: entry.workMinutes,
      workText: formatMinutesAsTimeText_(entry.workMinutes),
      breakMinutes: entry.breakMinutes,
      breakText: formatMinutesAsTimeText_(entry.breakMinutes),
      durationText: formatDurationText_(entry.workMinutes),
      records: entry.records
    })).sort((a, b) => (a.system || '').localeCompare(b.system || ''));

    const staffSummaries = Array.from(staffSummaryMap.values()).map(entry => ({
      staffType: entry.staffType,
      staffId: entry.staffId,
      staffName: entry.staffName,
      workMinutes: entry.workMinutes,
      workText: formatMinutesAsTimeText_(entry.workMinutes),
      breakMinutes: entry.breakMinutes,
      breakText: formatMinutesAsTimeText_(entry.breakMinutes),
      durationText: formatDurationText_(entry.workMinutes),
      workingDays: entry.workingDays,
      systems: Array.from(entry.systems).sort()
    })).sort((a, b) => {
      const typeDiff = (a.staffType || '').localeCompare(b.staffType || '');
      if (typeDiff !== 0) return typeDiff;
      return (a.staffName || '').localeCompare(b.staffName || '');
    });

    return {
      ok: true,
      dataset: {
        range: { from: range.fromKey, to: range.toKey },
        timezone: range.tz,
        totals: {
          workMinutes: totalWork,
          workText: formatMinutesAsTimeText_(totalWork),
          breakMinutes: totalBreak,
          breakText: formatMinutesAsTimeText_(totalBreak),
          durationText: formatDurationText_(totalWork)
        },
        systems: systemSummaries,
        staff: staffSummaries,
        records
      }
    };
  } catch (err) {
    const message = err && err.message ? err.message : String(err);
    Logger.log('[' + tag + '] ' + (err && err.stack ? err.stack : message));
    return { ok: false, message };
  }
}

function init_(){ ensureAuxSheets_(); }

/***** ログ・News *****/
function log_(op,pid,detail){
  sh('操作ログ').appendRow([new Date(), op, String(pid), detail||'', (Session.getActiveUser()||{}).getEmail()]);
}
function formatNewsRow_(pid, type, msg, meta){
  let metaStr = '';
  if (meta != null) {
    try {
      metaStr = typeof meta === 'string' ? meta : JSON.stringify(meta);
    } catch (e) {
      metaStr = String(meta);
    }
  }
  return [new Date(), String(pid), type, msg, '', metaStr, ''];
}

function setNewsClearedAt_(sheet, rowNumber){
  if (!sheet || !rowNumber) return;
  sheet.getRange(rowNumber, 5).setValue(new Date());
}

function getNewsDismissedColumn_(){
  const sheet = sh('News');
  const width = sheet.getLastColumn();
  const head = width > 0 ? sheet.getRange(1, 1, 1, width).getDisplayValues()[0] : [];
  const idx = head.findIndex(label => String(label || '').trim().toLowerCase() === 'dismissed');
  if (idx >= 0) return idx + 1;
  const newCol = width + 1;
  sheet.insertColumnsAfter(width, 1);
  sheet.getRange(1, newCol).setValue('dismissed');
  return newCol;
}

function parseNewsMetaValue_(value){
  if (value == null || value === '') return null;
  if (typeof value === 'object' && !Array.isArray(value)) return value;
  const text = String(value).trim();
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch (err) {
    return text;
  }
}

function resolveNewsMetaType_(meta){
  if (meta && typeof meta === 'object' && meta.type != null) {
    return String(meta.type);
  }
  if (typeof meta === 'string') {
    return String(meta);
  }
  return '';
}

function normalizeNewsMetaType_(metaType){
  const raw = String(metaType || '').trim();
  if (!raw) return '';
  switch (raw) {
    case 'consent_handout_followup':
    case 'consent_handout_follow-up':
    case 'consent_handout_followup_required':
    case 'consent_handout_pending':
      return 'consent_reminder';
    case 'consent_doctor_report':
    case 'consent_verification':
    case 'consent_verification_required':
      return 'consent_verification';
    case 'consent_handover':
    case 'consent_handover_pending':
    case 'handover':
      return 'handover';
    default:
      return raw;
  }
}

function normalizeConsentNewsMeta_(meta, message){
  const resolved = resolveNewsMetaType_(meta);
  let normalizedType = normalizeNewsMetaType_(resolved);
  if (!normalizedType) {
    const msg = String(message || '');
    if (msg.indexOf('同意書受渡が必要です') >= 0) {
      normalizedType = 'consent_reminder';
    } else if (msg.indexOf('同意期限50日前') >= 0) {
      normalizedType = 'consent_verification';
    } else if (msg.indexOf('受渡') >= 0 || msg.indexOf('受け渡し') >= 0) {
      normalizedType = 'handover';
    }
  }

  if (!normalizedType) {
    return { meta, metaType: '', changed: false };
  }

  let nextMeta = meta;
  let changed = false;
  if (!meta || typeof meta !== 'object') {
    nextMeta = { type: normalizedType };
    changed = true;
  } else if (meta.type !== normalizedType) {
    nextMeta = Object.assign({}, meta, { type: normalizedType });
    changed = true;
  } else if (normalizedType !== resolved) {
    changed = true;
  }

  return { meta: nextMeta, metaType: normalizedType, changed };
}

function isConsentReminderMessage_(message){
  if (!message) return false;
  const text = String(message);
  return (
    text.indexOf('同意書受渡が必要です') >= 0
      || text.indexOf('同意書の取得') >= 0
      || text.indexOf('同意書取得') >= 0
  );
}

function isConsentReminderNews_(news){
  if (!news) return false;
  const type = String(news.type || '').trim();
  if (type !== '同意') return false;
  const metaType = normalizeNewsMetaType_(resolveNewsMetaType_(news.meta));
  if (metaType === 'consent_reminder') return true;
  if (metaType === 'consent_verification') return false;
  return isConsentReminderMessage_(news.message);
}

function readNewsRows_(){
  const s = sh('News');
  const lr = s.getLastRow();
  if (lr < 2) return [];
  const width = Math.min(7, s.getLastColumn());
  const range = s.getRange(2, 1, lr - 1, width);
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const raw = values[i];
    const disp = displayValues[i];
    const rawDate = raw[0];
    let whenText = String(disp[0] || '').trim();
    if (!whenText && rawDate instanceof Date) {
      whenText = Utilities.formatDate(rawDate, timezone, 'yyyy-MM-dd HH:mm');
    }
    if (!whenText && rawDate != null && rawDate !== '') {
      whenText = String(rawDate);
    }
    const tsCandidate = rawDate instanceof Date
      ? rawDate.getTime()
      : (whenText ? new Date(whenText).getTime() : NaN);
    const rowNumber = 2 + i;
    const pidRaw = disp[1] != null && disp[1] !== '' ? disp[1] : raw[1];
    const normalizedPid = normId_(pidRaw);
    const typeText = String(disp[2] != null ? disp[2] : raw[2] || '');
    const messageText = String(disp[3] != null ? disp[3] : raw[3] || '');
    const metaRaw = width >= 6 ? raw[5] : '';
    let meta = parseNewsMetaValue_(metaRaw);
    if (typeText === '同意') {
      const normalizedMeta = normalizeConsentNewsMeta_(meta, messageText);
      meta = normalizedMeta.meta;
    }
    const clearedAtValue = width >= 5 ? raw[4] : '';
    const clearedAtText = String(disp[4] != null ? disp[4] : clearedAtValue || '').trim();
    let clearedAt = clearedAtText;
    if (!clearedAt && clearedAtValue instanceof Date) {
      clearedAt = Utilities.formatDate(clearedAtValue, timezone, 'yyyy-MM-dd HH:mm');
    }
    const dismissedRaw = width >= 7 ? raw[6] : '';
    const dismissed = toBooleanFromCell_(dismissedRaw);
    const ts = Number.isFinite(tsCandidate) ? tsCandidate : 0;
    rows.push({
      ts,
      when: whenText,
      rowNumber,
      pid: normalizedPid,
      type: typeText,
      message: messageText,
      meta,
      clearedAt,
      cleared: !!clearedAt,
      dismissed
    });
  }
  return rows;
}

function fetchNewsRowsForPid_(normalized){
  if (!normalized) return [];
  return readNewsRows_()
    .filter(row => !row.cleared && !row.dismissed && row.pid === normalized)
    .map(row => ({
      ts: row.ts,
      when: row.when,
      type: row.type,
      message: row.message,
      meta: row.meta,
      clearedAt: row.clearedAt,
      dismissed: !!row.dismissed,
      rowNumber: row.rowNumber,
      pid: row.pid
    }));
}

function fetchGlobalNewsRows_(){
  return readNewsRows_()
    .filter(row => !row.cleared && !row.dismissed && !row.pid)
    .map(row => ({
      ts: row.ts,
      when: row.when,
      type: row.type,
      message: row.message,
      meta: row.meta,
      clearedAt: row.clearedAt,
      dismissed: !!row.dismissed,
      rowNumber: row.rowNumber,
      pid: row.pid
    }));
}

function formatNewsOutput_(rows){
  if (!Array.isArray(rows) || !rows.length) return [];
  return rows
    .slice()
    .sort((a, b) => (b.ts || 0) - (a.ts || 0))
      .map(row => ({
        when: row.when,
        type: row.type,
        message: row.message,
        meta: row.meta,
        clearedAt: row.clearedAt,
        dismissed: !!row.dismissed,
        rowNumber: row.rowNumber,
        pid: row.pid
      }));
}

function pushNewsRows_(rows){
  if (!rows || !rows.length) return;
  let sheet;
  try {
    sheet = sh('News');
  } catch (err) {
    Logger.log('[pushNewsRows_] Failed to get News sheet: ' + (err && err.message ? err.message : err));
    try {
      ensureAuxSheets_({ force: true });
      sheet = sh('News');
    } catch (retryErr) {
      Logger.log('[pushNewsRows_] Retried ensureAuxSheets_ but still failed: ' + (retryErr && retryErr.message ? retryErr.message : retryErr));
      throw retryErr;
    }
  }
  const start = sheet.getLastRow() + 1;
  const width = rows[0].length;
  sheet.getRange(start, 1, rows.length, width).setValues(rows);
  Logger.log('[pushNewsRows_] appended rows: ' + rows.length);
  let hasGlobal = false;
  const affected = Array.from(new Set(rows.map(r => {
    const normalized = normId_(r && r[1]);
    if (!normalized) hasGlobal = true;
    return normalized;
  }).filter(Boolean)));
  if (affected.length) {
    invalidatePatientCaches_(affected, { news: true });
  }
  if (hasGlobal) {
    invalidateGlobalNewsCache_();
  }
}
function pushNews_(pid,type,msg,meta){
  pushNewsRows_([formatNewsRow_(pid, type, msg, meta)]);
}
function appendRowsToSheet_(sheetName, rows){
  if (!rows || !rows.length) return;
  const sheet = sh(sheetName);
  const width = rows[0].length;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, width).setValues(rows);
}
function getNews(pid){
  standardizeConsentNewsMeta_();
  const normalized = normId_(pid);
  const globalNews = cacheFetch_(GLOBAL_NEWS_CACHE_KEY, fetchGlobalNewsRows_, PATIENT_CACHE_TTL_SECONDS) || [];
  if (!normalized) {
    return formatNewsOutput_(globalNews);
  }
  const patientNews = cacheFetch_(PATIENT_CACHE_KEYS.news(normalized), () => fetchNewsRowsForPid_(normalized), PATIENT_CACHE_TTL_SECONDS) || [];
  return formatNewsOutput_(globalNews.concat(patientNews));
}
  function clearConsentRelatedNews_(pid){
    const s=sh('News'); const lr=s.getLastRow(); if(lr<2) return;
    const vals=s.getRange(2,1,lr-1,5).getValues(); // [TS,pid,type,msg,cleared]
    for (let i=0;i<vals.length;i++){
      if(String(vals[i][1])===String(pid)){
        const typ=String(vals[i][2]||'');
        const trimmed = typ.trim();
        if(typ.indexOf('同意')>=0 || typ.indexOf('期限')>=0 || typ.indexOf('予定')>=0 || trimmed === '再同意取得確認' || trimmed === '再同意'){
          setNewsClearedAt_(s, 2 + i);
        }
      }
    }
    invalidatePatientCaches_(pid, { news: true });
  }

function clearNewsByTypes_(pid, types){
  if(!Array.isArray(types) || !types.length) return;
  const normalized = types
    .map(t => String(t || '').trim())
    .filter(t => t.length);
  if(!normalized.length) return;
  const s = sh('News');
  const lr = s.getLastRow();
  if(lr < 2) return;
  const vals = s.getRange(2, 1, lr - 1, 5).getDisplayValues();
  const typeSet = new Set(normalized);
  for(let i=0;i<vals.length;i++){
    if(String(vals[i][1]) !== String(pid)) continue;
    const typ = String(vals[i][2] || '').trim();
    if(typeSet.has(typ)){
      setNewsClearedAt_(s, 2 + i);
    }
  }
  invalidatePatientCaches_(pid, { news: true });
}

function markNewsClearedByType(pid, type, options){
  const typeName = String(type || '').trim();
  if (!typeName) return 0;
  const s = sh('News');
  const lr = s.getLastRow();
  if (lr < 2) return 0;
  const width = Math.min(7, s.getMaxColumns());
  const vals = s.getRange(2, 1, lr - 1, width).getValues();
  const matchPid = String(pid || '').trim();
  const normalizedPid = normId_(matchPid);
  const filterMessage = options && options.messageContains ? String(options.messageContains) : '';
  const filterMetaType = options && options.metaType ? normalizeNewsMetaType_(options.metaType) : '';
  const filterRow = options && typeof options.rowNumber === 'number' ? Number(options.rowNumber) : null;
  const metaMatches = options && typeof options.metaMatches === 'object' && options.metaMatches
    ? options.metaMatches
    : null;
  const touchedPatients = new Set();
  let touchedGlobal = false;
  let cleared = 0;
  for (let i = 0; i < vals.length; i++) {
    const rowNumber = 2 + i;
    if (filterRow && rowNumber !== filterRow) continue;
    const rowPidRaw = vals[i][1];
    const rowPid = normId_(rowPidRaw);
    if (normalizedPid) {
      if (rowPid !== normalizedPid) continue;
    } else if (matchPid && String(rowPidRaw || '').trim() !== matchPid) {
      continue;
    }
    const rowType = String(vals[i][2] || '').trim();
    if (rowType !== typeName) continue;
    if (filterMessage) {
      const message = String(vals[i][3] || '');
      if (message.indexOf(filterMessage) < 0) continue;
    }
    let meta = null;
    if (filterMetaType || metaMatches) {
      const metaRaw = width >= 6 ? vals[i][5] : '';
      meta = parseNewsMetaValue_(metaRaw);
    }
    if (filterMetaType) {
      const resolvedType = normalizeNewsMetaType_(resolveNewsMetaType_(meta));
      if (resolvedType !== filterMetaType) continue;
    }
    if (metaMatches) {
      if (!meta || typeof meta !== 'object') continue;
      let metaOk = true;
      Object.keys(metaMatches).forEach(key => {
        if (!metaOk) return;
        const expected = metaMatches[key];
        const actual = meta[key];
        if (expected == null && actual == null) return;
        if (key === 'type') {
          if (normalizeNewsMetaType_(resolveNewsMetaType_(actual)) !== normalizeNewsMetaType_(expected)) {
            metaOk = false;
          }
          return;
        }
        if (String(actual) !== String(expected)) {
          metaOk = false;
        }
      });
      if (!metaOk) continue;
    }
    setNewsClearedAt_(s, rowNumber);
    cleared++;
    if (rowPid) {
      touchedPatients.add(rowPid);
    } else {
      touchedGlobal = true;
    }
  }
  if (cleared) {
    if (touchedPatients.size) {
      const ids = Array.from(touchedPatients);
      invalidatePatientCaches_(ids, { news: true });
    }
    if (touchedGlobal) {
      invalidateGlobalNewsCache_();
    }
  }
  return cleared;
}

function clearMonthlyHandoverReminder_(pid, monthKey){
  const matches = { type: 'handover_missing_monthly' };
  if (monthKey) {
    matches.month = monthKey;
  }
  return markNewsClearedByType(pid, '申し送り', {
    metaType: 'handover_missing_monthly',
    metaMatches: matches,
    messageContains: '申し送りが未入力'
  });
}

function clearDoctorReportMissingReminder_(pid, consentExpiry){
  const matches = { type: 'missing_moushiokuri' };
  if (consentExpiry != null && consentExpiry !== '') {
    matches.consentExpiry = String(consentExpiry);
  }
  return markNewsClearedByType(pid, '申し送り', {
    metaType: 'missing_moushiokuri',
    metaMatches: matches,
    messageContains: '申し送りが未入力'
  });
}

function clearNewsByTreatment_(treatmentId){
  if (!treatmentId) return;
  const s = sh('News');
  const lr = s.getLastRow();
  if (lr < 2) return;
  const width = Math.min(6, s.getMaxColumns());
  if (width < 5) return;
  const vals = s.getRange(2, 1, lr - 1, width).getValues();
  const metaIndex = width >= 6 ? 5 : -1;
  const clearedCol = 4; // 5列目（cleared）
  const matches = [];
  let touchedGlobal = false;
  for (let i = 0; i < vals.length; i++) {
    const metaText = metaIndex >= 0 ? String(vals[i][metaIndex] || '').trim() : '';
    if (!metaText) continue;
    let meta;
    try {
      meta = JSON.parse(metaText);
    } catch (e) {
      meta = null;
    }
    if (meta && String(meta.treatmentId || '') === String(treatmentId)) {
      matches.push(i);
      continue;
    }
    if (!meta && metaText === String(treatmentId)) {
      matches.push(i);
    }
  }
  const affected = new Set();
  matches.forEach(idx => {
    setNewsClearedAt_(s, 2 + idx);
    const pid = normId_(vals[idx][1]);
    if (pid) {
      affected.add(pid);
    } else {
      touchedGlobal = true;
    }
  });
  if (affected.size) {
    invalidatePatientCaches_(Array.from(affected), { news: true });
  }
  if (touchedGlobal) {
    invalidateGlobalNewsCache_();
  }
}

function buildLatestHandoverMap_(){
  const map = {};
  let sheet;
  try {
    sheet = ensureHandoverSheet_();
  } catch (e) {
    return map;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  values.forEach(row => {
    const pid = normId_(row[1]);
    if (!pid) return;
    const ts = row[0] instanceof Date
      ? row[0]
      : parseDateTimeFlexible_(row[0], tz) || parseDateFlexible_(row[0]);
    if (!(ts instanceof Date) || isNaN(ts.getTime())) return;
    const note = String(row[3] || '').trim();
    const time = ts.getTime();
    const existing = map[pid];
    if (!existing || time > existing.timestamp) {
      map[pid] = {
        timestamp: time,
        note,
        when: Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm')
      };
    }
  });
  return map;
}

function getLatestHandoverEntry_(pid, options){
  const normalized = normId_(pid);
  if (!normalized) return null;
  const map = options && options.map ? options.map : buildLatestHandoverMap_();
  return map[normalized] || null;
}

function isRecentHandoverEntry_(entry, referenceDate){
  if (!entry || !entry.note || !entry.timestamp) return false;
  const ref = referenceDate instanceof Date ? new Date(referenceDate.getTime()) : new Date();
  if (!(ref instanceof Date) || isNaN(ref.getTime())) return false;
  ref.setHours(0, 0, 0, 0);
  const entryDate = new Date(entry.timestamp);
  if (!(entryDate instanceof Date) || isNaN(entryDate.getTime())) return false;
  const diff = ref.getTime() - entryDate.getTime();
  if (diff < 0) return true;
  const windowMs = DOCTOR_REPORT_HANDOVER_WINDOW_DAYS * 24 * 60 * 60 * 1000;
  if (diff <= windowMs) return true;
  return entryDate.getFullYear() === ref.getFullYear() && entryDate.getMonth() === ref.getMonth();
}

function checkConsentExpiration_(){
  ensureAuxSheets_();
  const sheet = sh('患者情報');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, scanned: 0, inserted: 0 };
  const lastCol = sheet.getLastColumn();
  const head = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const cRec = getColFlexible_(head, LABELS.recNo, PATIENT_COLS_FIXED.recNo, '施術録番号');
  const cConsent = getColFlexible_(head, LABELS.consent, PATIENT_COLS_FIXED.consent, '同意年月日');
  if (!cRec || !cConsent) {
    return { ok: false, scanned: 0, inserted: 0, reason: 'missingColumns' };
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const today = new Date();
  const todayY = Number(Utilities.formatDate(today, tz, 'yyyy'));
  const todayM = Number(Utilities.formatDate(today, tz, 'MM')) - 1;
  const todayD = Number(Utilities.formatDate(today, tz, 'dd'));
  const todayStart = new Date(todayY, todayM, todayD);

  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const existing = readNewsRows_();
  const existingKeys = new Set();
  const existingDoctorReportKeys = new Set();
  const existingDoctorReportMissingKeys = new Set();
  existing.forEach(row => {
    if (row.cleared) return;
    if (!row.pid) return;
    const typeText = String(row.type || '').trim();
    const meta = row.meta;
    const metaType = normalizeNewsMetaType_(resolveNewsMetaType_(meta));
    let expiryKey = '';
    if (meta && typeof meta === 'object' && meta.consentExpiry) {
      expiryKey = String(meta.consentExpiry);
    }
    if (typeText === '申し送り') {
      if (meta && typeof meta === 'object' && meta.type === 'missing_moushiokuri') {
        existingDoctorReportMissingKeys.add(row.pid + '|' + expiryKey);
      }
      return;
    }
    if (typeText !== '同意') return;
    const message = String(row.message || '').trim();
    if (metaType === 'consent_reminder') {
      existingKeys.add(row.pid + '|' + expiryKey);
      return;
    }
    if (metaType === 'consent_verification') {
      existingDoctorReportKeys.add(row.pid + '|' + expiryKey);
      return;
    }
    if (message === '同意書受渡が必要です') {
      existingKeys.add(row.pid + '|' + expiryKey);
      return;
    }
    if (message.indexOf('同意期限50日前') >= 0) {
      existingDoctorReportKeys.add(row.pid + '|' + expiryKey);
    }
  });

  const toInsert = [];
  const insertedKeys = new Set();
  const insertedDoctorReportKeys = new Set();
  const insertedDoctorReportMissingKeys = new Set();
  const doctorReportRemindersToClear = new Map();
  const missingHandoverRemindersToClear = new Map();
  const latestHandoversMap = buildLatestHandoverMap_();
  const dayMs = 24 * 60 * 60 * 1000;
  const parseIsoLocal = (text) => {
    const m = text && text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  };

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const pidRaw = row[cRec - 1];
    const pidNormalized = normId_(pidRaw);
    if (!pidNormalized) continue;
    const consent = row[cConsent - 1];
    const expiryStr = calcConsentExpiry_(consent);
    if (!expiryStr) continue;
    const expiryDate = parseIsoLocal(expiryStr);
    if (!expiryDate) continue;
    expiryDate.setHours(0, 0, 0, 0);
    const reminderDate = new Date(expiryDate.getTime() - 30 * dayMs);
    reminderDate.setHours(0, 0, 0, 0);
    const daysFromReminder = Math.floor((todayStart.getTime() - reminderDate.getTime()) / dayMs);
    if (daysFromReminder < 0) continue; // 1か月前より未来の場合はスキップ
    const daysSinceExpiry = Math.floor((todayStart.getTime() - expiryDate.getTime()) / dayMs);
    if (daysSinceExpiry > 30) continue; // 期限を30日以上過ぎていたらスキップ
    const daysUntilExpiry = Math.floor((expiryDate.getTime() - todayStart.getTime()) / dayMs);
    const pidForNews = String(pidRaw || '').trim();
    if (!pidForNews) continue;
    const reportTriggerDate = new Date(expiryDate.getTime() - 50 * dayMs);
    reportTriggerDate.setHours(0, 0, 0, 0);
    const daysSinceReportTrigger = Math.floor((todayStart.getTime() - reportTriggerDate.getTime()) / dayMs);
    if (daysSinceReportTrigger >= 0 && daysUntilExpiry >= 0) {
      const reportKey = pidNormalized + '|' + expiryStr;
      const latestHandover = latestHandoversMap[pidNormalized] || null;
      const hasRecentHandover = isRecentHandoverEntry_(latestHandover, todayStart);
      if (!hasRecentHandover) {
        if (!existingDoctorReportMissingKeys.has(reportKey) && !insertedDoctorReportMissingKeys.has(reportKey)) {
          const missingMeta = {
            source: 'auto',
            type: 'missing_moushiokuri',
            consentExpiry: expiryStr,
            triggerDate: Utilities.formatDate(reportTriggerDate, tz, 'yyyy-MM-dd')
          };
          toInsert.push(formatNewsRow_(pidForNews, '申し送り', '申し送りが未入力のため報告書を生成できません。申し送りを入力してください。', missingMeta));
          insertedDoctorReportMissingKeys.add(reportKey);
        }
        if (existingDoctorReportKeys.has(reportKey) && !doctorReportRemindersToClear.has(reportKey)) {
          doctorReportRemindersToClear.set(reportKey, pidForNews);
        }
      } else {
        if (existingDoctorReportMissingKeys.has(reportKey) && !missingHandoverRemindersToClear.has(reportKey)) {
          missingHandoverRemindersToClear.set(reportKey, { pid: pidForNews, consentExpiry: expiryStr });
        }
        if (!existingDoctorReportKeys.has(reportKey) && !insertedDoctorReportKeys.has(reportKey)) {
          const reportMeta = {
            source: 'auto',
            type: 'consent_verification',
            consentExpiry: expiryStr,
            triggerDate: Utilities.formatDate(reportTriggerDate, tz, 'yyyy-MM-dd')
          };
          toInsert.push(formatNewsRow_(pidForNews, '同意', '⚠️ 同意期限50日前になりました', reportMeta));
          insertedDoctorReportKeys.add(reportKey);
        }
      }
    }
    const key = pidNormalized + '|' + expiryStr;
    if (existingKeys.has(key) || insertedKeys.has(key)) {
      continue;
    }
    const meta = {
      source: 'auto',
      type: 'consent_reminder',
      consentExpiry: expiryStr,
      reminderDate: Utilities.formatDate(reminderDate, tz, 'yyyy-MM-dd')
    };
    toInsert.push(formatNewsRow_(pidForNews, '同意', '同意書受渡が必要です', meta));
    insertedKeys.add(key);
  }

  if (doctorReportRemindersToClear.size) {
    doctorReportRemindersToClear.forEach(pidValue => {
      try {
        markNewsClearedByType(pidValue, '同意', {
          metaType: 'consent_verification',
          messageContains: '同意期限50日前'
        });
      } catch (err) {
        Logger.log('[checkConsentExpiration_] failed to clear doctor report reminder: ' + (err && err.message ? err.message : err));
      }
    });
  }
  if (missingHandoverRemindersToClear.size) {
    missingHandoverRemindersToClear.forEach(item => {
      try {
        clearDoctorReportMissingReminder_(item.pid, item.consentExpiry);
      } catch (err) {
        Logger.log('[checkConsentExpiration_] failed to clear missing handover reminder: ' + (err && err.message ? err.message : err));
      }
    });
  }
  if (toInsert.length) {
    pushNewsRows_(toInsert);
  }
  return { ok: true, scanned: rows.length, inserted: toInsert.length };
}

function checkConsentExpiration(){
  return checkConsentExpiration_();
}

function checkMonthlyHandovers_(){
  ensureAuxSheets_();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const today = new Date();
  const monthKey = Utilities.formatDate(today, tz, 'yyyy-MM');
  const monthStart = new Date(today.getFullYear(), today.getMonth(), 1, 0, 0, 0, 0);
  const monthEnd = new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59, 999);
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0, 0);

  const handoverSheet = ensureHandoverSheet_();
  const handoverSet = new Set();
  const handoverLastRow = handoverSheet.getLastRow();
  if (handoverLastRow >= 2) {
    const handoverValues = handoverSheet.getRange(2, 1, handoverLastRow - 1, 5).getValues();
    handoverValues.forEach(row => {
      const pid = normId_(row[1]);
      if (!pid) return;
      let ts = row[0];
      if (!(ts instanceof Date)) {
        ts = parseDateTimeFlexible_(ts, tz) || parseDateFlexible_(ts);
      }
      if (!(ts instanceof Date) || isNaN(ts.getTime())) return;
      const time = ts.getTime();
      if (time < monthStart.getTime() || time > monthEnd.getTime()) return;
      handoverSet.add(pid);
    });
  }

  const existingNews = readNewsRows_();
  const existingReminderKeys = new Set();
  existingNews.forEach(row => {
    if (row.cleared) return;
    if (!row.pid) return;
    if (String(row.type || '').trim() !== '申し送り') return;
    const meta = row.meta;
    if (meta && typeof meta === 'object' && meta.type === 'handover_missing_monthly') {
      if (!meta.month || meta.month === monthKey) {
        existingReminderKeys.add(row.pid);
      }
      return;
    }
    const message = String(row.message || '');
    if (message.indexOf('申し送りが未入力') >= 0) {
      existingReminderKeys.add(row.pid);
    }
  });

  const statusMap = buildPatientStatusMap_();
  const patientSheet = sh('患者情報');
  const lastRow = patientSheet.getLastRow();
  if (lastRow < 2) {
    return { ok: true, month: monthKey, scanned: 0, inserted: 0 };
  }
  const lastCol = patientSheet.getLastColumn();
  const head = patientSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const cRec = getColFlexible_(head, LABELS.recNo, PATIENT_COLS_FIXED.recNo, '施術録番号');
  const rows = patientSheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const reminders = [];
  let scanned = 0;

  rows.forEach(row => {
    const pidRaw = row[cRec - 1];
    const pidNormalized = normId_(pidRaw);
    if (!pidNormalized) return;
    scanned += 1;
    if (handoverSet.has(pidNormalized)) return;
    if (existingReminderKeys.has(pidNormalized)) return;
    const statusInfo = statusMap[pidNormalized] || { status: 'active', pauseUntil: '' };
    if (statusInfo.status === 'stopped') return;
    if (statusInfo.status === 'suspended') {
      const pauseUntil = parseDateFlexible_(statusInfo.pauseUntil);
      if (pauseUntil && pauseUntil.getTime() >= todayStart.getTime()) {
        return;
      }
    }
    const pidForNews = String(pidRaw || '').trim() || pidNormalized;
    const meta = { type: 'handover_missing_monthly', month: monthKey };
    reminders.push(formatNewsRow_(pidForNews, '申し送り', '今月の申し送りが未入力です', meta));
  });

  if (reminders.length) {
    pushNewsRows_(reminders);
  }

  return { ok: true, month: monthKey, scanned, inserted: reminders.length };
}

function checkMonthlyHandovers(){
  return checkMonthlyHandovers_();
}

/***** ステータス（休止/中止） *****/
function buildPatientStatusMap_(){
  const map = {};
  let sheet;
  try {
    sheet = sh('フラグ');
  } catch (e) {
    return map;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;
  const values = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  values.forEach(row => {
    const pid = normId_(row[0]);
    if (!pid) return;
    map[pid] = {
      status: row[1] || 'active',
      pauseUntil: row[2] || ''
    };
  });
  return map;
}

function getStatus_(pid){
  const s=sh('フラグ'); const lr=s.getLastRow(); if (lr<2) return {status:'active', pauseUntil:''};
  const vals=s.getRange(2,1,lr-1,3).getDisplayValues();
  const row=vals.reverse().find(r=> String(r[0])===String(pid));
  if (!row) return {status:'active', pauseUntil:''};
  return { status: row[1]||'active', pauseUntil: row[2]||'' };
}
function markSuspend(pid){
  ensureAuxSheets_();
  const until = Utilities.formatDate(new Date(Date.now()+1000*60*60*24*30), Session.getScriptTimeZone()||'Asia/Tokyo','yyyy-MM-dd');
  sh('フラグ').appendRow([String(pid),'suspended',until]);
  pushNews_((pid),'状態','休止に設定（ミュート '+until+' まで）');
  log_('休止', pid, until);
  invalidatePatientCaches_(pid, { header: true });
}
function markStop(pid){
  ensureAuxSheets_();
  sh('フラグ').appendRow([String(pid),'stopped','']);
  pushNews_(pid,'状態','中止に設定（以降のリマインド停止）');
  log_('中止', pid, '');
  invalidatePatientCaches_(pid, { header: true });
}

/***** ヘッダ正規化ユーティリティ *****/
function normalizeHeaderKey_(s){
  if(!s) return '';
  const z2h = String(s).normalize('NFKC');
  const noSpace = z2h.replace(/\s+/g,'');
  const noPunct = noSpace.replace(/[（）\(\)\[\]【】:：・\-＿_]/g,'');
  return noPunct.toLowerCase();
}
function buildHeaderMap_(headersRow){
  const map={};
  headersRow.forEach((h,i)=>{
    const k=normalizeHeaderKey_(h);
    if(k && !map[k]) map[k]=i+1;
  });
  return map;
}
function resolveColByLabels_(headersRow, labelCandidates, fieldLabel, required=true){
  const idx=buildHeaderMap_(headersRow);
  for(const label of labelCandidates){
    const k=normalizeHeaderKey_(label);
    if(idx[k]) return idx[k];
  }
  if(required) throw new Error('患者情報に見出しが見つかりません: '+fieldLabel+'（候補: '+labelCandidates.join('/')+'）');
  return null;
}
function getColFlexible_(headersRow, labelCandidates, fallback1Based, fieldLabel){
  const c = resolveColByLabels_(headersRow, labelCandidates, fieldLabel, false);
  return c || fallback1Based;
}

/***** ID正規化（"0007" ≒ "7" を同一視） *****/
function normId_(x){
  if (x == null) return '';
  let s = String(x).normalize('NFKC').replace(/\s+/g,'');
  s = s.replace(/^0+/, '');
  return s;
}

/***** 患者行の安全取得（見出しの揺れに耐える） *****/
function findPatientRow_(pid){
  const pnorm = normId_(pid);
  const s = sh('患者情報');
  const lr = s.getLastRow(); if (lr < 2) return null;
  const lc = s.getLastColumn();
  const head = s.getRange(1,1,1,lc).getDisplayValues()[0];
  const cRec = getColFlexible_(head, LABELS.recNo, PATIENT_COLS_FIXED.recNo, '施術録番号');
  const vals = s.getRange(2,1,lr-1,lc).getValues();
  for (let i=0; i<vals.length; i++){
    const v = normId_(vals[i][cRec-1]);
    if (v && v === pnorm){
      return {
        row: 2+i, lc, head,
        rowValues: s.getRange(2+i, 1, 1, lc).getDisplayValues()[0]
      };
    }
  }
  return null;
}

function findLatestTreatmentRow_(pid){
  const normalized = normId_(pid);
  if (!normalized) return null;
  const cacheKey = PATIENT_CACHE_KEYS.latestTreatmentRow(normalized);
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        const parsed = JSON.parse(cached);
        if (parsed && typeof parsed === 'object') {
          return parsed;
        }
      } catch (e) {}
    }
  } catch (err) {
    Logger.log('[findLatestTreatmentRow_] cache fetch skipped: ' + (err && err.message ? err.message : err));
  }
  const sheet = sh('施術録');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const maxCols = sheet.getMaxColumns();
  const width = Math.min(TREATMENT_SHEET_HEADER.length, maxCols);
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  let latestIndex = -1;
  let latestTs = -Infinity;
  values.forEach((row, idx) => {
    const pidCell = normId_(row[1]);
    if (pidCell !== normalized) return;
    const tsVal = row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime();
    if (!Number.isFinite(tsVal)) return;
    if (latestIndex < 0 || tsVal > latestTs) {
      latestIndex = idx;
      latestTs = tsVal;
    }
  });
  if (latestIndex < 0) return null;
  const rowNumber = latestIndex + 2;
  const latest = {
    row: rowNumber,
    head: headers,
    rowValues: sheet.getRange(rowNumber, 1, 1, width).getDisplayValues()[0],
    rawValues: values[latestIndex],
    width
  };
  try {
    const cache = CacheService.getScriptCache();
    cache.put(cacheKey, JSON.stringify(latest), PATIENT_CACHE_TTL_SECONDS);
  } catch (err) {
    Logger.log('[findLatestTreatmentRow_] cache put skipped: ' + (err && err.message ? err.message : err));
  }
  return latest;
}

/***** 負担割合 正規化 *****/
function normalizeBurdenRatio_(text) {
  if (!text) return null;
  const t = String(text).replace(/\s/g,'').replace('％','%').replace('割','');
  if (/^[123]$/.test(t)) return Number(t)/10;                 // 1,2,3
  if (/^(10|20|30)%?$/.test(t)) return Number(RegExp.$1)/100; // 10/20/30 or 10%
  return null;
}
function toBurdenDisp_(ratio) {
  if (ratio === 0.1) return '1割';
  if (ratio === 0.2) return '2割';
  if (ratio === 0.3) return '3割';
  return '';
}
/** 入力（1割/2/20% など）→ { num:1|2|3|null, disp:'1割|2割|3割|'' } */
function parseShareToNumAndDisp_(text){
  const r = normalizeBurdenRatio_(text); // 0.1 / 0.2 / 0.3 or null
  if (r === 0.1) return { num: 1, disp: '1割' };
  if (r === 0.2) return { num: 2, disp: '2割' };
  if (r === 0.3) return { num: 3, disp: '3割' };
  return { num: null, disp: '' };
}
/***** 日付パース（和暦・略号対応）＆ 同意期限 *****/
function parseDateFlexible_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const raw = String(v).trim();
  if (!raw) return null;

  // 和暦（正式）
  const era = raw.match(/(令和|平成|昭和)\s*(\d+)[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
  if (era) {
    const eraName = era[1], y = Number(era[2]), m = Number(era[3]), d = Number(era[4]);
    const base = eraName === '令和' ? 2018 : eraName === '平成' ? 1988 : 1925; // R1=2019, H1=1989, S1=1926
    return new Date(base + y, m - 1, d);
  }
  // 和暦（略号 R/H/S）
  const eraShort = raw.match(/([RrHhSs])\s*(\d+)[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
  if (eraShort) {
    const ch = eraShort[1].toUpperCase(), y = Number(eraShort[2]), m = Number(eraShort[3]), d = Number(eraShort[4]);
    const base = ch === 'R' ? 2018 : ch === 'H' ? 1988 : 1925;
    return new Date(base + y, m - 1, d);
  }
  // 西暦
  const m1 = raw.match(/(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
  if (m1) return new Date(Number(m1[1]), Number(m1[2]) - 1, Number(m1[3]));
  // yyyymmdd
  const n = raw.replace(/\D/g,'');
  if (n.length === 8) return new Date(Number(n.slice(0,4)), Number(n.slice(4,6))-1, Number(n.slice(6,8)));

  const d = new Date(raw);
  return isNaN(d.getTime()) ? null : d;
}
function calcConsentExpiry_(consentVal) {
  const d = parseDateFlexible_(consentVal);
  if (!d) return '';
  const day = d.getDate();
  const base = new Date(d);
  // 1〜15日 → +5か月の月末 / 16日〜 → +6か月の月末
  if (day <= 15) base.setMonth(base.getMonth() + 5, 1);
  else           base.setMonth(base.getMonth() + 6, 1);
  const end = new Date(base.getFullYear(), base.getMonth() + 1, 0);
  return Utilities.formatDate(end, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd');
}

/***** 月次・直近 *****/
function getMonthlySummary_(pid) {
  const s = sh('施術録'); const lr = s.getLastRow();
  if (lr < 2) return { current:{count:0,est:0}, previous:{count:0,est:0} };
  const vals = s.getRange(2,1,lr-1,6).getValues();
  const now = new Date();
  const first=(y,m)=>new Date(y,m,1);
  const last=(y,m)=>new Date(y,m+1,0,23,59,59);
  const y=now.getFullYear(), m=now.getMonth();
  const curS=first(y,m), curE=last(y,m);
  const prevS=first(y,m-1), prevE=last(y,m-1);
  let c=0,p=0;
  vals.forEach(r=>{
    const ts=r[0], id=String(r[1]);
    if (id!==String(pid)) return;
    const d = ts instanceof Date ? ts : new Date(ts);
    if (isNaN(d.getTime())) return;
    if (d>=curS && d<=curE) c++; else if (d>=prevS && d<=prevE) p++;
  });
  const unit = APP.BASE_FEE_YEN || 4170;
  return { current:{count:c, est: Math.round(c*unit*0.1)}, previous:{count:p, est: Math.round(p*unit*0.1)} };
}
function getRecentActivity_(pid) {
  const s=sh('施術録'); const lr=s.getLastRow();
  let lastTreat='';
  if (lr>=2) {
    const v=s.getRange(2,1,lr-1,6).getValues().filter(r=> String(r[1])===String(pid));
    if (v.length) {
      const d=v[v.length-1][0];
      const dd = d instanceof Date ? d : new Date(d);
      if (!isNaN(dd.getTime())) lastTreat = Utilities.formatDate(dd, Session.getScriptTimeZone()||'Asia/Tokyo','yyyy-MM-dd');
    }
  }
  const sp=sh('患者情報'); const lc=sp.getLastColumn();
  const head=sp.getRange(1,1,1,lc).getDisplayValues()[0];
  const cRec = getColFlexible_(head, LABELS.recNo,  PATIENT_COLS_FIXED.recNo,  '施術録番号');
  const cCons= getColFlexible_(head, LABELS.consent,PATIENT_COLS_FIXED.consent,'同意年月日');
  let lastConsent='';
  const vals=sp.getRange(2,1,sp.getLastRow()-1,lc).getDisplayValues();
  const row=vals.find(r=> String(r[cRec-1])===String(pid));
  lastConsent = row ? (row[cCons-1]||'') : '';
  return { lastTreat, lastConsent, lastStaff: '' };
}

/***** 患者ヘッダ（画面表示用） *****/
function getPatientHeader(pid){
  const normalized = normId_(pid);
  if (!normalized) return null;
  const cacheKey = PATIENT_CACHE_KEYS.header(normalized);
  try {
    SpreadsheetApp.flush();
    Utilities.sleep(60);
  } catch (err) {
    console.warn('[getPatientHeader] cache bypass failed', err);
  }
  return cacheFetch_(cacheKey, () => {
    ensureAuxSheets_();
    const hit = findPatientRow_(pid);
    if (!hit) return null;

    const s = sh('患者情報'), head = hit.head, rowV = hit.rowValues;
    const cName = getColFlexible_(head, LABELS.name,     PATIENT_COLS_FIXED.name,     '名前');
    const cHos  = getColFlexible_(head, LABELS.hospital, PATIENT_COLS_FIXED.hospital, '病院名');
    const cDoc  = getColFlexible_(head, LABELS.doctor,   PATIENT_COLS_FIXED.doctor,   '医師');
    const cFuri = getColFlexible_(head, LABELS.furigana, PATIENT_COLS_FIXED.furigana, 'ﾌﾘｶﾞﾅ');
    const cBirth= getColFlexible_(head, LABELS.birth,    PATIENT_COLS_FIXED.birth,    '生年月日');
    const cCons = getColFlexible_(head, LABELS.consent,  PATIENT_COLS_FIXED.consent,  '同意年月日');
    const cConsHandout = getColFlexible_(head, LABELS.consentHandout, PATIENT_COLS_FIXED.consentHandout, '配布');
    const cShare= getColFlexible_(head, LABELS.share,    PATIENT_COLS_FIXED.share,    '負担割合');
    const cTel  = getColFlexible_(head, LABELS.phone,    PATIENT_COLS_FIXED.phone,    '電話');
    const cConsentContent = getColFlexible_(head, LABELS.consentContent, PATIENT_COLS_FIXED.consentContent, '同意症状');

    // 年齢
    const bd = parseDateFlexible_(rowV[cBirth-1]||'');
    let age=null, ageClass='';
    if (bd) {
      const t=new Date();
      age = t.getFullYear()-bd.getFullYear() - ((t.getMonth()<bd.getMonth() || (t.getMonth()===bd.getMonth() && t.getDate()<bd.getDate()))?1:0);
      if (age>=75) ageClass='後期高齢'; else if (age>=65) ageClass='前期高齢';
    }

    // 同意期限
    const consent = rowV[cCons-1]||'';
    const consentHandout = rowV[cConsHandout-1]||'';
    const expiry  = calcConsentExpiry_(consent) || '—';

    // 負担割合
    const shareRaw  = rowV[cShare-1]||'';
    const shareNorm = normalizeBurdenRatio_(shareRaw);
    const shareDisp = shareNorm ? toBurdenDisp_(shareNorm) : shareRaw;

    const monthly = getMonthlySummary_(pid);
    const recent  = getRecentActivity_(pid);
    const stat    = getStatus_(pid);

    const header = {
      patientId:String(normId_(pid)),
      name: rowV[cName-1]||'',
      furigana: rowV[cFuri-1]||'',
      hospital: rowV[cHos-1]||'',
      doctor:   rowV[cDoc-1]||'',
      phone:    rowV[cTel-1]||'',
      birth:    rowV[cBirth-1]||'',
      age, ageClass,
      consentDate: consent || '',
      consentHandoutDate: consentHandout || '',
      consentExpiry: expiry,
      consentContent: cConsentContent ? String(rowV[cConsentContent-1] || '').trim() : '',
      burden: shareDisp || '',
      monthly, recent,
      status: stat.status,
      pauseUntil: stat.pauseUntil
    };

    try {
      console.log('[getPatientHeader]', String(normId_(pid)), JSON.stringify(header));
    } catch (err) {
      console.warn('[getPatientHeader] log failed', err);
    }

    return header;
  }, PATIENT_CACHE_TTL_SECONDS);
}

function getPatientBundle(pid){
  const normalized = normId_(pid);
  if (!normalized) {
    return { header: null, news: [], treatments: [] };
  }

  const header = getPatientHeader(normalized);
  const news = (getNews(normalized) || []).map(item => Object.assign({}, item, {
    htmlMessage: convertPlainTextToSafeHtml_(item && item.message ? item.message : '')
  }));
  const treatments = listTreatmentsForCurrentMonth(normalized);

  return { header, news, treatments };
}

/***** ID候補 *****/
function listPatientIds(){
  const s=sh('患者情報'); const lr=s.getLastRow(); if(lr<2) return [];
  const lc=s.getLastColumn(); const head=s.getRange(1,1,1,lc).getDisplayValues()[0];
  const cRec = getColFlexible_(head, LABELS.recNo, PATIENT_COLS_FIXED.recNo, '施術録番号');
  const cName = getColFlexible_(head, LABELS.name, PATIENT_COLS_FIXED.name, '名前');
  const cFuri = getColFlexible_(head, LABELS.furigana, PATIENT_COLS_FIXED.furigana, 'ﾌﾘｶﾞﾅ');
  const vals=s.getRange(2,1,lr-1,lc).getDisplayValues();
  const seen = new Set();
  const out = [];
  vals.forEach(r=>{
    const id = normId_(r[cRec-1]);
    if(!id || seen.has(id)) return;
    seen.add(id);
    out.push({
      id,
      name: r[cName-1] || '',
      kana: (cFuri && r[cFuri-1]) ? r[cFuri-1] : ''
    });
  });
  return out;
}

/***** 定型文 *****/
function getPresets(){
  ensureAuxSheets_();
  const s = sh('定型文'); const lr = s.getLastRow();
  if (lr < 2) {
    return [
      {cat:'所見',label:'特記事項なし',text:'特記事項なし。経過良好。'},
      {cat:'所見',label:'バイタル安定',text:'バイタル安定。生活指導継続。'},
      {cat:'所見',label:'請求書・領収書受渡',text:'請求書・領収書を受け渡し済み。'},
      {cat:'所見',label:'配布物受渡',text:'配布物（説明資料）を受け渡し済み。'},
      {cat:'所見',label:'同意書受渡',text:'同意書受渡。'},
      {cat:'所見',label:'再同意取得確認',text:'再同意の取得を確認。引き続き施術を継続。'}
    ];
  }
  const vals = s.getRange(2,1,lr-1,3).getDisplayValues(); // [カテゴリ, ラベル, 文章]
  return vals
    .map(r=>({cat:r[0],label:r[1],text:r[2]}))
    .filter(preset => String(preset && preset.label || '').trim() !== '同意書取得確認');
}

/***** 施術保存 *****/
function queueAfterTreatmentJob(job){
  if (!job || typeof job !== 'object') return;

  const key = 'AFTER_JOBS';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('[queueAfterTreatmentJob] Failed to acquire lock');
    scheduleAfterTreatmentJobTrigger_({ force: true });
    return;
  }
  try {
    const p = PropertiesService.getScriptProperties();
    let jobs = [];
    try {
      const raw = p.getProperty(key);
      jobs = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(jobs)) jobs = [];
    } catch (e) {
      Logger.log('[queueAfterTreatmentJob] Failed to parse existing jobs: ' + (e && e.message ? e.message : e));
      jobs = [];
    }
    jobs.push(job);
    p.setProperty(key, JSON.stringify(jobs));
    Logger.log('[queueAfterTreatmentJob] Queued: ' + JSON.stringify(job));
  } finally {
    lock.releaseLock();
  }

  const processedInline = drainAfterTreatmentJobs_({ inline: true });
  if (!processedInline) {
    scheduleAfterTreatmentJobTrigger_();
  }
}

function afterTreatmentJob(){
  drainAfterTreatmentJobs_({ triggered: true });
}

function drainAfterTreatmentJobs_(options){
  const key = 'AFTER_JOBS';
  const inline = options && options.inline;
  let lock;
  try {
    lock = LockService.getScriptLock();
  } catch (err) {
    Logger.log('[afterTreatmentJob] LockService unavailable: ' + (err && err.message ? err.message : err));
    return false;
  }
  const waitMs = options && typeof options.waitMs === 'number'
    ? options.waitMs
    : (inline ? 50 : 5000);
  let gotLock = false;
  try {
    gotLock = lock.tryLock(waitMs);
  } catch (err) {
    Logger.log('[afterTreatmentJob] Failed to acquire lock: ' + (err && err.message ? err.message : err));
    gotLock = false;
  }
  if (!gotLock) {
    if (!inline) {
      Logger.log('[afterTreatmentJob] Failed to acquire lock');
    }
    return false;
  }

  let jobs = [];
  try {
    const p = PropertiesService.getScriptProperties();
    const raw = p.getProperty(key);
    p.deleteProperty(key);
    p.deleteProperty(AFTER_TREATMENT_TRIGGER_KEY);
    if (raw) {
      try {
        jobs = JSON.parse(raw) || [];
        if (!Array.isArray(jobs)) jobs = [];
      } catch (e) {
        Logger.log('[afterTreatmentJob] Failed to parse jobs: ' + (e && e.message ? e.message : e));
        jobs = [];
      }
    }
  } finally {
    try { lock.releaseLock(); } catch (err) {
      Logger.log('[afterTreatmentJob] Failed to release lock: ' + (err && err.message ? err.message : err));
    }
  }

  if (!jobs.length) {
    return false;
  }

  Logger.log('[afterTreatmentJob] Executing jobs: ' + jobs.length + (inline ? ' (inline)' : ''));
  executeAfterTreatmentJobs_(jobs);
  return true;
}

function executeAfterTreatmentJobs_(jobs){
  if (!Array.isArray(jobs) || !jobs.length) return;
  ensureAuxSheets_();
  const newsRows = [];
  const scheduleRows = [];
  const userEmail = (Session.getActiveUser()||{}).getEmail() || '';
  const tz = Session.getScriptTimeZone()||'Asia/Tokyo';

  jobs.forEach(job=>{
    try {
      const pid = job.patientId;
      const treatmentMeta = job.treatmentId ? { source: 'treatment', treatmentId: job.treatmentId } : null;
      const addNews = (type, message, extraMeta) => {
        let meta = null;
        if (treatmentMeta) {
          meta = Object.assign({}, treatmentMeta);
        }
        if (extraMeta) {
          meta = meta ? Object.assign(meta, extraMeta) : Object.assign({}, extraMeta);
        }
        newsRows.push(formatNewsRow_(pid, type, message, meta));
      };

      // News / 同意日 / 負担割合 / 予定登録など重い処理をここでまとめて実行
      let consentReminderPushed = false;
      if (job.presetLabel){
        if (job.presetLabel.indexOf('同意書受渡') >= 0){
          if (job.consentUndecided){
            addNews('同意','通院日未定です。後日確認してください。', { type: 'consent_reminder', reason: 'consent_undecided' });
            consentReminderPushed = true;
          } else {
            const visitPlanDate = job.visitPlanDate ? String(job.visitPlanDate).trim() : '';
            const followupMessageBase = '通院日が近づいています。ご利用者様に声かけをしてください。';
            const followupMessage = visitPlanDate
              ? `${followupMessageBase}（通院予定：${visitPlanDate}）`
              : followupMessageBase;
            const meta = { type: 'consent_reminder', reason: 'handout_followup' };
            if (visitPlanDate) {
              meta.visitPlanDate = visitPlanDate;
            }
            addNews('同意', followupMessage, meta);
          }
        }
      }
      if (job.consentUndecided && !consentReminderPushed){
        addNews('同意','通院日未定です。後日確認してください。', { type: 'consent_reminder', reason: 'consent_undecided' });
      }
      if (job.burdenShare){
        updateBurdenShare(pid, job.burdenShare, treatmentMeta ? { meta: treatmentMeta } : undefined);
      }
      if (job.visitPlanDate){
        scheduleRows.push([String(pid),'通院', job.visitPlanDate, userEmail]);
        addNews('予定','通院予定を登録：' + job.visitPlanDate);
      }
      log_('施術後処理', pid, JSON.stringify(job));
    } catch (e) {
      Logger.log('[afterTreatmentJob] Job failed: ' + (e && e.message ? e.message : e));
    }
  });

  if (scheduleRows.length) {
    appendRowsToSheet_('予定', scheduleRows);
  }
  if (newsRows.length) {
    pushNewsRows_(newsRows);
  }
}


/***** 当月の施術一覧 取得・更新・削除 *****/
function listTreatmentsForCurrentMonth(pid){
  const normalized = normId_(pid);
  if (!normalized) return [];
  const cacheKey = PATIENT_CACHE_KEYS.treatments(normalized);
  return cacheFetch_(cacheKey, () => {
    const s = sh('施術録');
    const lr = s.getLastRow();
    if (lr < 2) return [];
    const rows = lr - 1;
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0);
    const end   = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
    const timestamps = s.getRange(2, 1, rows, 1).getValues();
    const ids = s.getRange(2, 2, rows, 1).getDisplayValues();
    const notes = s.getRange(2, 3, rows, 1).getValues();
    const emails = s.getRange(2, 4, rows, 1).getValues();
    const treatmentIdRange = s.getRange(2, 7, rows, 1);
    const treatmentIds = treatmentIdRange.getValues();
    const categories = s.getRange(2, 8, rows, 1).getValues();
    const missingTreatmentIds = [];

    const out = [];
    for (let i = 0; i < rows; i++) {
      const pidCell = normId_(ids[i][0]);
      if (pidCell !== normalized) continue;
      const ts = timestamps[i][0];
      const d = ts instanceof Date ? ts : new Date(ts);
      const timestamp = d instanceof Date ? d.getTime() : NaN;
      if (!Number.isFinite(timestamp)) continue;
      if (d < start || d > end) continue;
      let treatmentId = String((treatmentIds[i] && treatmentIds[i][0]) || '').trim();
      if (!treatmentId) {
        treatmentId = Utilities.getUuid();
        treatmentIds[i][0] = treatmentId;
        missingTreatmentIds.push(i);
      }
      const categoryLabel = String((categories[i] && categories[i][0]) || '');
      const categoryKey = mapTreatmentCategoryCellToKey_(categoryLabel);
      out.push({
        row: 2 + i,
        when: Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm'),
        note: String((notes[i] && notes[i][0]) || ''),
        email: String((emails[i] && emails[i][0]) || ''),
        treatmentId,
        category: categoryLabel,
        categoryKey,
        timestamp
      });
    }

    if (missingTreatmentIds.length) {
      try {
        treatmentIdRange.setValues(treatmentIds);
      } catch (err) {
        Logger.log('[listTreatmentsForCurrentMonth] failed to backfill treatmentId: ' + (err && err.message ? err.message : err));
      }
    }

    return out
      .sort((a, b) => {
        const aTs = Number.isFinite(a.timestamp) ? a.timestamp : 0;
        const bTs = Number.isFinite(b.timestamp) ? b.timestamp : 0;
        if (aTs !== bTs) return bTs - aTs;
        return b.row - a.row;
      })
      .map(row => {
        const { timestamp, ...rest } = row;
        return rest;
      });
  }, PATIENT_CACHE_TTL_SECONDS);
}
function updateTreatmentRow(row, note) {
  const s = sh('施術録');
  if (row <= 1 || row > s.getLastRow()) throw new Error('行が不正です');

  const newNote = String(note || '').trim();

  // 直前の値を取得
  const oldNote = String(s.getRange(row, 3).getValue() || '').trim();
  const pid = String(s.getRange(row, 2).getValue() || '').trim();

  // 🔒 二重編集チェック
  if (oldNote === newNote) {
    return { ok: false, skipped: true, msg: '変更内容が直前と同じのため編集をスキップしました' };
  }

  // 書き換え
  s.getRange(row, 3).setValue(newNote);

  // ログ
  log_('施術修正', '(row:' + row + ')', newNote);

  if (pid) {
    invalidatePatientCaches_(pid, { header: true, treatments: true, latestTreatmentRow: true });
  }

  return { ok: true, updatedRow: row, newNote };
}

function deleteTreatmentRow(treatmentId){
  const s = sh('施術録');
  const normalizedTreatmentId = String(treatmentId || '').trim();
  if (!normalizedTreatmentId) throw new Error('削除対象の施術IDが指定されていません');

  const found = findTreatmentRowById_(s, normalizedTreatmentId);
  if (!found || !found.rowNumber) throw new Error('指定した施術記録が見つかりませんでした');

  const targetRow = found.rowNumber;
  const rowVals = (found && found.row && Array.isArray(found.row))
    ? found.row
    : s.getRange(targetRow, 1, 1, Math.min(TREATMENT_SHEET_HEADER.length, s.getMaxColumns())).getValues()[0];
  const pid = String(rowVals[1] || '').trim();

  s.deleteRow(targetRow);
  clearNewsByTreatment_(normalizedTreatmentId);
  log_('施術削除', `(row:${targetRow})`, '');
  if (pid) {
    invalidatePatientCaches_(pid, { header: true, treatments: true, latestTreatmentRow: true });
  }
  try {
    const normalizedPid = normId_(pid);
    const treatments = normalizedPid ? listTreatmentsForCurrentMonth(normalizedPid) : [];
    return { ok: true, patientId: normalizedPid, treatments };
  } catch (err) {
    Logger.log('[deleteTreatmentRow] failed to fetch updated list: ' + (err && err.message ? err.message : err));
    return { ok: true, patientId: normId_(pid) || '', treatments: null };
  }
}

function splitTreatmentNoteForSummary_(text){
  const lines = String(text || '').split(/\n+/).map(s => s.trim()).filter(Boolean);
  const vitals = lines.filter(line => /^vital\b/i.test(line));
  const others = lines.filter(line => !/^vital\b/i.test(line));
  return {
    note: others.join(' '),
    vitals: vitals.join(' '),
    raw: lines.join(' ')
  };
}

function getTreatmentNotesInRange_(pid, startDate, endDate){
  const s = sh('施術録');
  const lr = s.getLastRow();
  if (lr < 2) return [];
  const vals = s.getRange(2, 1, lr - 1, 4).getValues();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const target = String(normId_(pid));
  const startTs = startDate instanceof Date ? startDate.getTime() : null;
  const endTs = endDate instanceof Date ? endDate.getTime() : null;
  const out = [];

  vals.forEach(row => {
    const rowPid = String(normId_(row[1]));
    if (rowPid !== target) return;
    const ts = row[0] instanceof Date ? row[0] : parseDateTimeFlexible_(row[0], tz) || parseDateFlexible_(row[0]);
    if (!(ts instanceof Date) || isNaN(ts.getTime())) return;
    const ms = ts.getTime();
    if (startTs != null && ms < startTs) return;
    if (endTs != null && ms > endTs) return;
    const when = Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm');
    const parts = splitTreatmentNoteForSummary_(String(row[2] || ''));
    out.push({ when, note: parts.note, vitals: parts.vitals, raw: parts.raw, timestamp: ms });
  });

  out.sort((a, b) => a.timestamp - b.timestamp);
  return out.map(item => ({ when: item.when, note: item.note, vitals: item.vitals, raw: item.raw }));
}

function getHandoversInRange_(pid, startDate, endDate){
  const s = ensureHandoverSheet_();
  const lr = s.getLastRow();
  if (lr < 2) return [];
  const vals = s.getRange(2, 1, lr - 1, 5).getValues();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const target = String(normId_(pid));
  const startTs = startDate instanceof Date ? startDate.getTime() : null;
  const endTs = endDate instanceof Date ? endDate.getTime() : null;
  const out = [];

  vals.forEach(row => {
    const rowPid = String(normId_(row[1]));
    if (rowPid !== target) return;
    const ts = row[0] instanceof Date ? row[0] : parseDateTimeFlexible_(row[0], tz) || parseDateFlexible_(row[0]);
    if (!(ts instanceof Date) || isNaN(ts.getTime())) return;
    const ms = ts.getTime();
    if (startTs != null && ms < startTs) return;
    if (endTs != null && ms > endTs) return;
    const when = Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm');
    const note = String(row[3] || '').trim();
    out.push({ when, note, timestamp: ms });
  });

  out.sort((a, b) => a.timestamp - b.timestamp);
  return out.map(item => ({ when: item.when, note: item.note }));
}

function resolveIcfSummaryRange_(rangeKey){
  const normalized = normalizeRangeInputObject_(rangeKey);
  const key = normalized && normalized.key ? normalized.key : 'all';
  const now = new Date();
  const defaultEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  let end = defaultEnd;
  let start = null;
  let label = '全期間';
  let monthsValue = 'all';
  const formatYmd = (date) => {
    if (!(date instanceof Date) || isNaN(date.getTime())) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  };

  switch (key) {
    case '1m':
    case '2m':
    case '3m':
    case '6m':
    case '12m': {
      const months = Number(String(key).replace('m', '')) || 0;
      monthsValue = String(months || '');
      label = `直近${months}か月`;
      start = new Date(end.getTime());
      start.setHours(0, 0, 0, 0);
      start.setMonth(start.getMonth() - months);
      break;
    }
    case 'custom': {
      const startCandidate = parseDateFlexible_(normalized.start || '');
      const endCandidate = parseDateFlexible_(normalized.end || '') || parseDateFlexible_(normalized.start || '');
      if (!startCandidate || !endCandidate) {
        throw new Error('カスタム期間の開始日と終了日を指定してください。');
      }
      const startDate = new Date(startCandidate.getFullYear(), startCandidate.getMonth(), startCandidate.getDate(), 0, 0, 0, 0);
      const endDate = new Date(endCandidate.getFullYear(), endCandidate.getMonth(), endCandidate.getDate(), 23, 59, 59, 999);
      if (startDate.getTime() > endDate.getTime()) {
        throw new Error('カスタム期間の開始日が終了日より後になっています。');
      }
      start = startDate;
      end = endDate;
      monthsValue = normalized.months != null ? String(normalized.months) : 'custom';
      const endLabelDate = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate(), 0, 0, 0, 0);
      label = `${formatYmd(startDate)}〜${formatYmd(endLabelDate)}`;
      break;
    }
    case 'all':
    default:
      label = '全期間';
      start = null;
      monthsValue = 'all';
      break;
  }

  if (start) {
    start = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 0, 0, 0, 0);
  }

  return {
    key,
    startDate: start,
    endDate: end,
    label,
    months: monthsValue,
    customStart: normalized.start || '',
    customEnd: normalized.end || ''
  };
}

/***** 同意・負担割合 更新（findPatientRow_ベース） *****/
function updateConsentDate(pid, dateStr, options){
  const hit = findPatientRow_(pid);
  if (!hit) throw new Error('患者が見つかりません');
  const s=sh('患者情報'); const head=hit.head;
  const cCons= getColFlexible_(head, LABELS.consent, PATIENT_COLS_FIXED.consent, '同意年月日');
  const cHandout = getColFlexible_(head, LABELS.consentHandout, PATIENT_COLS_FIXED.consentHandout, '配布');
  const metaRaw = options && options.meta ? options.meta : null;
  const metaType = normalizeNewsMetaType_(resolveNewsMetaType_(metaRaw));
  let meta = metaRaw;
  if (!metaType) {
    meta = meta && typeof meta === 'object' ? Object.assign({}, meta) : {};
    meta.type = 'consent_verification';
  } else if (meta && typeof meta === 'object') {
    meta = Object.assign({}, meta, { type: metaType });
  } else {
    meta = metaType;
  }
  const source = meta && meta.source ? String(meta.source) : '';
  const isTreatmentTriggered = source === 'treatment';

  if (isTreatmentTriggered) {
    s.getRange(hit.row, cHandout).setValue(dateStr || '');
  } else {
    s.getRange(hit.row, cCons).setValue(dateStr || '');
  }

  clearConsentRelatedNews_(pid);

  const newsMessage = dateStr
    ? '再同意取得確認（同意日更新：' + dateStr + '）'
    : '再同意取得確認（同意日更新）';
  pushNews_(pid,'同意', newsMessage, meta);

  const logDetail = isTreatmentTriggered ? '確認日:' + (dateStr || '') : (dateStr || '');
  log_('同意日更新', pid, logDetail);
  invalidatePatientCaches_(pid, { header: true });
}

function dismissConsentReminder(payload){
  const pidRaw = payload && payload.patientId ? payload.patientId : payload;
  const pid = normId_(pidRaw);
  if (!pid) throw new Error('患者IDが指定されていません');
  const newsRow = payload && typeof payload.newsRow === 'number' ? Number(payload.newsRow) : null;
  const dismissedCol = getNewsDismissedColumn_();
  const sheet = sh('News');
  const targetRow = (() => {
    if (newsRow && newsRow >= 2) return newsRow;
    const rows = readNewsRows_();
    const match = rows.find(row => row.pid === pid && isConsentReminderNews_(row) && !row.dismissed && !row.cleared);
    return match ? match.rowNumber : null;
  })();
  if (!targetRow) {
    throw new Error('対象のお知らせが見つかりません');
  }
  sheet.getRange(targetRow, dismissedCol).setValue(true);
  invalidatePatientCaches_(pid, { news: true });
  invalidateGlobalNewsCache_();
  return { ok: true, rowNumber: targetRow };
}

function dismissHandoverReminder(payload){
  const pidRaw = payload && payload.patientId ? payload.patientId : payload;
  const pid = normId_(pidRaw);
  if (!pid) throw new Error('患者IDが指定されていません');
  const newsRow = payload && typeof payload.newsRow === 'number' ? Number(payload.newsRow) : null;
  const newsType = String(payload && payload.newsType || '').trim();
  const newsMessage = String(payload && payload.newsMessage || '');
  const newsMetaType = payload && payload.newsMetaType ? normalizeNewsMetaType_(payload.newsMetaType) : '';

  const dismissedCol = getNewsDismissedColumn_();
  const sheet = sh('News');
  const targetRow = (() => {
    if (newsRow && newsRow >= 2) return newsRow;
    const rows = readNewsRows_();
    const match = rows.find(row => {
      if (row.pid !== pid) return false;
      if (row.cleared || row.dismissed) return false;
      if (newsType && String(row.type || '').trim() !== newsType) return false;
      const rowMetaType = normalizeNewsMetaType_(row.meta);
      if (newsMetaType && rowMetaType && rowMetaType !== newsMetaType) return false;
      if (newsMessage && String(row.message || '').indexOf(newsMessage) < 0) return false;
      return true;
    });
    return match ? match.rowNumber : null;
  })();

  if (!targetRow) {
    throw new Error('対象のお知らせが見つかりません');
  }

  sheet.getRange(targetRow, dismissedCol).setValue(true);
  invalidatePatientCaches_(pid, { news: true });
  invalidateGlobalNewsCache_();
  return { ok: true, rowNumber: targetRow };
}
function updateBurdenShare(pid, shareText, options){
  const hit = findPatientRow_(pid);
  if (!hit) throw new Error('患者が見つかりません');
  const s=sh('患者情報'); const headers=hit.head;

  // 書き込み先列（患者情報の「負担割合」列）
  const cShare= getColFlexible_(headers, LABELS.share, PATIENT_COLS_FIXED.share, '負担割合');

  // 1) 入力を正規化 → num(1/2/3) と disp('1割/2割/3割')
  const parsed = parseShareToNumAndDisp_(shareText);

  // 2) 患者情報には数値で保存（例：2）※ null の場合は元の文字列をそのまま保存
  if (parsed.num != null) {
    s.getRange(hit.row, cShare).setValue(parsed.num); // ← 数値 1|2|3 を保存
  } else {
    s.getRange(hit.row, cShare).setValue(shareText || '');
  }

  // 3) 代表へ通知＆News
  const disp = parsed.disp || String(shareText||'');
  const meta = options && options.meta ? options.meta : null;
  pushNews_(pid,'通知','負担割合を更新：' + disp, meta);
  log_('負担割合更新', pid, disp);

  // 4) 施術録にも記録を残す（監査・検索用）
  const user = (Session.getActiveUser()||{}).getEmail();
  sh('施術録').appendRow([new Date(), String(pid), '負担割合を更新：' + (disp || shareText || ''), user, '', '', Utilities.getUuid() ]);

  invalidatePatientCaches_(pid, { header: true, treatments: true, latestTreatmentRow: true });
  return true;
}


/***** 請求集計（回数/負担/請求額） *****/
function parseBillingMonth_(text) {
  const trimmed = String(text || '').trim();
  const match = trimmed.match(/^(\d{4})-(0[1-9]|1[0-2])$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  return {
    year,
    month,
    ym: match[1] + '-' + match[2],
    sheetSuffix: match[1] + match[2]
  };
}

/***** PDF保存（Doc→PDFエクスポート方式：確実にPDF化） *****/
function getParentFolder_(){
  const id = (APP.PARENT_FOLDER_ID || PropertiesService.getScriptProperties().getProperty('PARENT_FOLDER_ID') || '').trim();
  if (id) return DriveApp.getFolderById(id);
  const file = DriveApp.getFileById(ss().getId());
  const it = file.getParents();
  if (it.hasNext()) return it.next();
  return DriveApp.getRootFolder();
}

function getPayrollPdfRootFolder_(){
  const props = PropertiesService.getScriptProperties();
  const configured = (props.getProperty(PAYROLL_PDF_ROOT_FOLDER_PROPERTY_KEY) || '').trim();
  if (configured) {
    try {
      return DriveApp.getFolderById(configured);
    } catch (err) {
      Logger.log('[getPayrollPdfRootFolder_] invalid property: ' + err);
    }
  }
  const fallback = (APP.PAYROLL_PDF_ROOT_FOLDER_ID || '').trim();
  if (fallback) {
    try {
      return DriveApp.getFolderById(fallback);
    } catch (err2) {
      Logger.log('[getPayrollPdfRootFolder_] invalid APP fallback: ' + err2);
    }
  }
  return getParentFolder_();
}

function formatPayrollEmployeeFolderName_(name){
  const raw = String(name || '').trim();
  const normalized = raw ? raw.replace(/\s+/g, '　') : '未設定従業員';
  return normalized.endsWith('殿') ? normalized : normalized + '殿';
}

function ensurePayrollEmployeeFolder_(employeeName){
  const root = getPayrollPdfRootFolder_();
  if (!root) throw new Error('給与明細の保存先フォルダを取得できません。');
  const existing = findPayrollEmployeeFolder_(employeeName);
  if (existing) return existing;
  return root.createFolder(formatPayrollEmployeeFolderName_(employeeName));
}

function findPayrollEmployeeFolder_(employeeName){
  const root = getPayrollPdfRootFolder_();
  if (!root) return null;
  const folderName = formatPayrollEmployeeFolderName_(employeeName);
  const iterator = root.getFoldersByName(folderName);
  if (iterator.hasNext()) {
    return iterator.next();
  }
  return null;
}

function deletePayrollPayslipFilesByName_(folder, fileName){
  if (!folder || !fileName) return;
  const iterator = folder.getFilesByName(fileName);
  while (iterator.hasNext()) {
    const file = iterator.next();
    try {
      file.setTrashed(true);
    } catch (err) {
      Logger.log('[deletePayrollPayslipFilesByName_] Failed to delete %s: %s', fileName, err);
    }
  }
}

function extractPayrollMonthHint_(fileName){
  const text = String(fileName || '');
  const match = text.match(/(\d{4})年(\d{1,2})月/);
  if (match) {
    return match[1] + '年' + Number(match[2]) + '月';
  }
  return '';
}

function listPayrollPayslipFilesInFolder_(folder, options){
  if (!folder) return [];
  const limit = Number(options && options.limit) > 0 ? Number(options.limit) : 50;
  const tz = (options && options.tz) || getConfig('timezone') || 'Asia/Tokyo';
  const iterator = folder.getFiles();
  const entries = [];
  while (iterator.hasNext()) {
    const file = iterator.next();
    const name = file.getName();
    const mime = file.getMimeType();
    const isPdf = (mime && mime.toLowerCase().indexOf('pdf') !== -1) || String(name || '').toLowerCase().endsWith('.pdf');
    if (!isPdf) continue;
    entries.push({
      id: file.getId(),
      name,
      url: file.getUrl(),
      createdAt: file.getDateCreated(),
      updatedAt: file.getLastUpdated(),
      sizeBytes: file.getSize()
    });
  }
  entries.sort((a, b) => {
    const aTime = a.updatedAt instanceof Date ? a.updatedAt.getTime() : 0;
    const bTime = b.updatedAt instanceof Date ? b.updatedAt.getTime() : 0;
    return bTime - aTime;
  });
  return entries.slice(0, limit).map(entry => ({
    id: entry.id,
    name: entry.name,
    url: entry.url,
    createdAt: entry.createdAt ? entry.createdAt.toISOString() : null,
    createdAtText: entry.createdAt ? formatIsoStringWithOffset_(entry.createdAt, tz) : '',
    updatedAt: entry.updatedAt ? entry.updatedAt.toISOString() : null,
    updatedAtText: entry.updatedAt ? formatIsoStringWithOffset_(entry.updatedAt, tz) : '',
    sizeBytes: entry.sizeBytes,
    sizeText: formatFileSizeShort_(entry.sizeBytes),
    monthHint: extractPayrollMonthHint_(entry.name)
  }));
}
function getOrCreateFolderForPatientMonth_(pid, date){
  const parent = getParentFolder_();
  const ym = Utilities.formatDate(date, Session.getScriptTimeZone()||'Asia/Tokyo', 'yyyy年M月');
  const it1 = parent.getFoldersByName(ym); const m = it1.hasNext()? it1.next() : parent.createFolder(ym);
  const it2 = m.getFoldersByName(String(pid)); return it2.hasNext()? it2.next() : m.createFolder(String(pid));
}
function savePdf_(pid, title, body){
  const folder = getOrCreateFolderForPatientMonth_(pid, new Date());

  // 一時Doc作成
  const doc = DocumentApp.create(title.replace(/\.pdf$/i,''));
  const docId = doc.getId();
  const dBody = doc.getBody();
  dBody.clear();
  body.split('\n').forEach(line => dBody.appendParagraph(line));
  doc.saveAndClose();

  // PDFにエクスポート
  const url = 'https://www.googleapis.com/drive/v3/files/'+docId+'/export?mimeType=application%2Fpdf';
  const token = ScriptApp.getOAuthToken();
  const pdfBlob = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  }).getBlob().setName(title);

  const file = folder.createFile(pdfBlob);

  // 索引記録
  sh('添付索引').appendRow([new Date(), String(pid),
    Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'Asia/Tokyo','yyyy-MM'),
    file.getName(), file.getId(), 'pdf', (Session.getActiveUser()||{}).getEmail()
  ]);
  pushNews_(pid,'PDF作成', file.getName()+' を作成しました');
  log_('PDF作成', pid, title);

  // 一時Doc削除（不要なら残してOK）
  DriveApp.getFileById(docId).setTrashed(true);

  return { ok:true, fileId:file.getId(), name:file.getName() };
}

function ensureChildFolder_(parent, name){
  if (!parent || !name) return null;
  const trimmed = String(name).trim();
  if (!trimmed) return parent;
  const iterator = parent.getFoldersByName(trimmed);
  if (iterator.hasNext()) {
    return iterator.next();
  }
  return parent.createFolder(trimmed);
}

function normalizeDoctorSectionHeading_(raw){
  const plain = String(raw || '')
    .replace(/[【】\s]/g, '')
    .replace(/[：:]/g, '')
    .trim();
  if (!plain) return '';
  const normalized = plain.replace(/・/g, '');
  const map = {
    '施術の内容頻度': '施術の内容・頻度',
    '施術内容頻度': '施術の内容・頻度',
    '施術内容': '施術の内容・頻度',
    '施術の内容': '施術の内容・頻度',
    '施術': '施術の内容・頻度',
    '施術頻度': '施術頻度',
    '患者の状態経過': '患者の状態・経過',
    '患者状態経過': '患者の状態・経過',
    '患者の状態': '患者の状態・経過',
    '患者状態': '患者の状態・経過',
    '患者経過': '患者の状態・経過',
    '状態経過': '患者の状態・経過',
    '経過': '患者の状態・経過',
    '報告内容': '報告内容',
    '特記すべき事項': '特記すべき事項',
    '特記事項': '特記すべき事項',
    '同意内容': '同意内容',
    '今後の方針': '今後の方針'
  };
  return map[normalized] || '';
}

function parseDoctorReportTextSections_(text){
  const lines = String(text || '').split(/\r?\n/);
  const map = {};
  let current = '';
  const setCurrent = (rawHeading, rest) => {
    const normalized = normalizeDoctorSectionHeading_(rawHeading);
    if (!normalized) return false;
    current = normalized;
    if (!map[current]) map[current] = [];
    const tail = rest != null ? String(rest).trim() : '';
    if (tail) {
      map[current].push(tail);
    }
    return true;
  };

  lines.forEach(raw => {
    const line = String(raw != null ? raw : '');
    const trimmed = line.trim();
    if (!trimmed) {
      if (current && map[current]) {
        map[current].push('');
      }
      return;
    }

    const bracket = trimmed.match(/^【([^】]+)】\s*(.*)$/);
    if (bracket && setCurrent(bracket[1], bracket[2])) {
      return;
    }

    const generic = trimmed.match(/^(?:[■□◆◇▶▷▶︎▸▹▶️➡＞>\-\*\s]*)([^：:】]+?)(?:\s*[：:]\s*(.*))?$/);
    if (generic && setCurrent(generic[1], generic[2])) {
      return;
    }

    if (!current) {
      return;
    }
    if (!map[current]) {
      map[current] = [];
    }
    map[current].push(trimmed);
  });

  const normalized = {};
  Object.keys(map).forEach(key => {
    const segments = [];
    let previousBlank = true;
    map[key].forEach(part => {
      const textPart = String(part != null ? part : '');
      const trimmed = textPart.trim();
      if (!trimmed) {
        if (!previousBlank && segments.length) {
          segments.push('');
          previousBlank = true;
        }
        return;
      }
      segments.push(trimmed);
      previousBlank = false;
    });
    const joined = segments.join('\n').trim();
    if (joined) {
      normalized[key] = joined;
    }
  });
  return normalized;
}

function normalizeDoctorReportTextForStorage_(text){
  const sections = parseDoctorReportTextSections_(text);
  const select = (...keys) => {
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      const value = sections[key];
      if (value && String(value).trim()) {
        return String(value).trim();
      }
    }
    return '';
  };

  const section1 = select('施術の内容・頻度', '施術内容');
  const section2 = select('患者の状態・経過', '報告内容');
  const rawSection3 = select('特記すべき事項', '特記事項');
  const section3 = rawSection3 || '特記すべき事項はありません。';
  const frequencyText = select('施術頻度');

  const blocks = [];
  const addBlock = (label, value) => {
    const trimmed = String(value || '').trim();
    if (!trimmed) return;
    if (blocks.length) blocks.push('');
    blocks.push(`【${label}】`);
    blocks.push(trimmed);
  };

  if (section1) addBlock('施術の内容・頻度', section1);
  if (section2) addBlock('患者の状態・経過', section2);
  addBlock('特記すべき事項', section3);

  const normalizedText = blocks.join('\n').trim();
  const resultText = normalizedText || String(text || '').trim();

  return {
    text: resultText,
    section1,
    section2,
    section3,
    frequencyText
  };
}

function buildDoctorReportPdfData_(patientId){
  const header = getPatientHeader(patientId);
  if (!header) {
    return { ok: false, code: 'patient_not_found', message: '患者情報が見つかりません。' };
  }

  const history = fetchReportHistoryForPid_(header.patientId);
  const entry = Array.isArray(history)
    ? history.find(item => item && item.audience === 'doctor')
    : null;
  if (!entry) {
    return { ok: false, code: 'report_not_found', message: '医師向け報告書が保存されていません。' };
  }

  const sections = parseDoctorReportTextSections_(entry.text || '');
  const consent = (sections['同意内容'] && String(sections['同意内容']).trim())
    || getConsentContentForPatient_(header.patientId)
    || '';
  const section1 = sections['施術の内容・頻度'] || sections['施術内容'] || '';
  const section2 = sections['患者の状態・経過'] || sections['報告内容'] || '';
  const section3 = sections['特記すべき事項'] || sections['特記事項'] || '';
  const frequencySource = sections['施術頻度'] || '';
  const frequencyText = frequencySource && String(frequencySource).trim()
    ? String(frequencySource).trim()
    : determineTreatmentFrequencyLabel_(countTreatmentsInRecentMonth_(header.patientId, new Date()));

  const treatmentLines = [];
  if (consent && String(consent).trim()) {
    treatmentLines.push('同意内容：' + String(consent).trim());
  }
  if (frequencyText) {
    treatmentLines.push('施術頻度：' + frequencyText);
  }
  const section1Text = section1 && String(section1).trim();
  if (section1Text) {
    treatmentLines.push(section1Text);
  }
  const treatmentSummary = treatmentLines.length ? treatmentLines.join('\n') : '施術頻度：情報不足';

  const reportSummary = section2 && String(section2).trim() ? String(section2).trim() : String(entry.text || '').trim();
  const closingSentence = '今後も安全に配慮しながら施術を継続してまいります。';
  let plan = sections['今後の方針'] && String(sections['今後の方針']).trim() || '';
  if (!plan) {
    if (reportSummary.indexOf(closingSentence) >= 0) {
      plan = closingSentence;
    } else if (section3 && String(section3).indexOf(closingSentence) >= 0) {
      plan = closingSentence;
    } else {
      plan = closingSentence;
    }
  }

  let remarks = section3 && String(section3).trim() ? String(section3).trim() : '';
  if (!remarks) {
    const specialList = Array.isArray(entry.special) ? entry.special.filter(Boolean) : [];
    if (specialList.length) {
      remarks = specialList.join('\n');
    }
  }
  if (!remarks) {
    remarks = '特記すべき事項はありません。';
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const todayText = Utilities.formatDate(new Date(), tz, 'yyyy年MM月dd日');

  return {
    ok: true,
    patientId: header.patientId,
    rangeLabel: entry.rangeLabel || '',
    data: {
      hospitalName: header.hospital || '',
      doctorName: header.doctor || '',
      patientName: header.name || '',
      birthDate: header.birth || '',
      consentText: consent || '',
      frequencyText: frequencyText || '',
      section1: section1Text || '',
      section2: reportSummary || '',
      section3: remarks || '',
      treatmentSummary,
      reportSummary,
      plan,
      remarks,
      createdDate: todayText
    }
  };
}

function createDoctorReportPdfFile_(prepared){
  if (!prepared || !prepared.data) {
    throw new Error('PDF生成に必要な情報が不足しています。');
  }
  if (!APP.DOCTOR_REPORT_TEMPLATE_ID || !APP.DOCTOR_REPORT_ROOT_FOLDER_ID) {
    throw new Error('医師向け報告書のテンプレートまたは保存先が設定されていません。');
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const todayText = prepared.data.createdDate || Utilities.formatDate(new Date(), tz, 'yyyy年MM月dd日');
  const root = DriveApp.getFolderById(APP.DOCTOR_REPORT_ROOT_FOLDER_ID);
  const pdfRoot = ensureChildFolder_(root, '報告書PDF');
  if (!pdfRoot) {
    throw new Error('報告書PDFフォルダを取得できません。');
  }
  const doctorFolder = ensureChildFolder_(pdfRoot, '医師');
  if (!doctorFolder) {
    throw new Error('医師向け報告書の保存先を取得できません。');
  }
  const baseName = `医師報告書_${prepared.data.patientName || prepared.patientId || '不明'}_${todayText}`;
  const template = DriveApp.getFileById(APP.DOCTOR_REPORT_TEMPLATE_ID);
  const copy = template.makeCopy(baseName, doctorFolder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  const replacements = {
    '{{病院名}}': prepared.data.hospitalName || '',
    '{{医師}}': prepared.data.doctorName || '',
    '{{患者名}}': prepared.data.patientName || '',
    '{{生年月日}}': prepared.data.birthDate || '',
    '{{同意内容}}': prepared.data.consentText || '',
    '{{施術頻度}}': prepared.data.frequencyText || '',
    '{{施術内容}}': prepared.data.section1 || '',
    '{{施術の内容・頻度}}': prepared.data.treatmentSummary || '',
    '{{報告内容}}': prepared.data.reportSummary || '',
    '{{患者経過}}': prepared.data.section2 || '',
    '{{患者の状態・経過}}': prepared.data.section2 || '',
    '{{今後の方針}}': prepared.data.plan || '',
    '{{特記事項}}': prepared.data.remarks || '',
    '{{特記すべき事項}}': prepared.data.remarks || '',
    '{{作成日}}': todayText
  };
  Object.keys(replacements).forEach(key => {
    try {
      body.replaceText(key, replacements[key]);
    } catch (err) {
      Logger.log(`[createDoctorReportPdfFile_] replace failed for ${key}: ` + (err && err.message ? err.message : err));
    }
  });
  doc.saveAndClose();

  const pdfBlob = copy.getAs(MimeType.PDF);
  const pdfName = baseName + '.pdf';
  pdfBlob.setName(pdfName);
  const pdfFile = doctorFolder.createFile(pdfBlob);
  copy.setTrashed(true);

  const createdAt = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  try {
    sh('添付索引').appendRow([
      new Date(),
      String(prepared.patientId || ''),
      Utilities.formatDate(new Date(), tz, 'yyyy-MM'),
      pdfFile.getName(),
      pdfFile.getId(),
      'pdf',
      (Session.getActiveUser() || {}).getEmail()
    ]);
  } catch (indexErr) {
    Logger.log('[createDoctorReportPdfFile_] failed to append index: ' + (indexErr && indexErr.message ? indexErr.message : indexErr));
  }

  return {
    file: pdfFile,
    createdAt
  };
}

function generateDoctorReportPdf(payload){
  assertDomain_();
  const idInput = payload && (payload.patientId || payload.pid || payload.id || payload.patientID);
  const patientId = normId_(idInput);
  if (!patientId) {
    throw new Error('患者IDが指定されていません。');
  }

  const prepared = buildDoctorReportPdfData_(patientId);
  if (!prepared.ok) {
    return {
      ok: false,
      code: prepared.code,
      message: prepared.message
    };
  }

  const result = createDoctorReportPdfFile_(prepared);
  const file = result.file;
  return {
    ok: true,
    patientId: prepared.patientId,
    rangeLabel: prepared.rangeLabel,
    fileId: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
    createdAt: result.createdAt
  };
}

/***** 文章整形（OpenAI → ローカルフォールバック） *****/
function getOpenAiKey_(){
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  return key ? key.trim() : '';
}
function extractSentencesForIcf_(text){
  return String(text || '')
    .split(/[。\.\!\?\n]+/)
    .map(s => s.trim())
    .filter(Boolean);
}
function countTreatmentsInRecentMonth_(pid, untilDate){
  const end = untilDate instanceof Date ? new Date(untilDate.getTime()) : new Date();
  end.setHours(23, 59, 59, 999);
  const start = new Date(end.getTime());
  start.setMonth(start.getMonth() - 1);
  start.setHours(0, 0, 0, 0);
  const notes = getTreatmentNotesInRange_(pid, start, end);
  return Array.isArray(notes) ? notes.length : 0;
}

function determineTreatmentFrequencyLabel_(count){
  const n = isFinite(count) ? Math.max(0, Math.round(count)) : 0;
  let label = '情報不足';
  if (n > 0 && n < 4) label = '週1回';
  else if (n >= 4 && n < 8) label = '週2回';
  else if (n >= 8 && n < 15) label = '週3回';
  else if (n >= 15) label = '週4回以上';
  return `${label}（直近1か月 ${n}回）`;
}

function getConsentContentForPatient_(pid){
  try {
    const hit = findPatientRow_(pid);
    if (!hit) return '';
    const { head, rowValues } = hit;
    const cConsentContent = getColFlexible_(head, LABELS.consentContent, PATIENT_COLS_FIXED.consentContent, '同意症状');
    if (!cConsentContent) return '';
    return String(rowValues[cConsentContent - 1] || '').trim();
  } catch (err) {
    Logger.log('getConsentContentForPatient_ error: ' + err);
    return '';
  }
}


function normalizeDoctorReportText_(text){
  return String(text || '')
    .replace(/\s*\n+\s*/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

function ensureDoctorSentenceWithFallback_(text, fallback){
  const normalize = (value) => normalizeDoctorReportText_(value);
  const ensurePeriod = (value) => {
    const norm = normalize(value);
    if (!norm) return '';
    return /[。．！？!？]$/.test(norm) ? norm : norm + '。';
  };
  const primary = ensurePeriod(text);
  if (primary) return primary;
  return ensurePeriod(fallback);
}

function parseDoctorSpecialList_(value){
  if (Array.isArray(value)) {
    return value
      .map(v => normalizeDoctorReportText_(v))
      .filter(Boolean);
  }
  if (value && typeof value === 'object') {
    if (Array.isArray(value.special)) {
      return value.special
        .map(v => normalizeDoctorReportText_(v))
        .filter(Boolean);
    }
    return [];
  }
  const raw = normalizeDoctorReportText_(value);
  if (!raw) return [];
  if (/^\[.*\]$/.test(raw)) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        return parsed
          .map(v => normalizeDoctorReportText_(v))
          .filter(Boolean);
      }
    } catch (e) {
      return [];
    }
  }
  return raw
    .split(/[,、\n]+/)
    .map(v => normalizeDoctorReportText_(v))
    .filter(Boolean);
}

function normalizeDoctorSpecialList_(value){
  const unique = Array.from(new Set(parseDoctorSpecialList_(value)));
  return unique.length ? unique : ['特記すべき事項はありません。'];
}

function buildDoctorStatusFromSections_(sections){
  const base = { body: '', activities: '', participation: '', environment: '', safety: '', special: [] };
  const priority = { body: -Infinity, activities: -Infinity, participation: -Infinity, environment: -Infinity, safety: -Infinity, special: -Infinity };
  const assignField = (key, value, score) => {
    if (!Object.prototype.hasOwnProperty.call(base, key)) return;
    const norm = normalizeDoctorReportText_(value);
    if (!norm) return;
    if (score < priority[key]) return;
    if (score === priority[key] && base[key]) return;
    base[key] = norm;
    priority[key] = score;
  };
  const assignSpecial = (value, score) => {
    const list = normalizeDoctorSpecialList_(value);
    if (!list.length) return;
    if (score < priority.special) return;
    if (score === priority.special && base.special.length) return;
    base.special = list;
    priority.special = score;
  };
  const mergeObject = (obj, score = 0) => {
    if (!obj || typeof obj !== 'object') return;
    if (obj.status && typeof obj.status === 'object') {
      mergeObject(obj.status, score);
    }
    assignField('body', obj.body, score);
    assignField('activities', obj.activities, score);
    assignField('participation', obj.participation, score);
    assignField('environment', obj.environment, score);
    assignField('safety', obj.safety, score);
    if (obj.special != null) assignSpecial(obj.special, score);
    if (!obj.body && obj.general) assignField('body', obj.general, score - 1);
  };

  if (sections && typeof sections === 'object' && !Array.isArray(sections)) {
    mergeObject(sections, 0);
  }

  if (Array.isArray(sections)) {
    sections.forEach(section => {
      const key = String(section && section.key ? section.key : '').toLowerCase();
      const data = section && typeof section.data === 'object' ? section.data : null;
      if (data) {
        mergeObject(data, 5);
      }
      if (!key) return;
      if (key === 'doctor_json' || key === 'doctor_status' || key === 'doctor_status_json') {
        const raw = section && section.json != null ? section.json : section && section.value != null ? section.value : section && section.text;
        if (raw && typeof raw === 'string') {
          try {
            const parsed = JSON.parse(raw);
            mergeObject(parsed, 10);
          } catch (e) {
            // ignore parse errors
          }
        } else if (raw && typeof raw === 'object') {
          mergeObject(raw, 10);
        }
        return;
      }
      if (Object.prototype.hasOwnProperty.call(base, key)) {
        assignField(key, section && section.text != null ? section.text : section && section.value, 1);
        return;
      }
      if (key === 'special') {
        const rawSpecial = data && data.special != null ? data.special : section && section.value != null ? section.value : section && section.text;
        assignSpecial(rawSpecial, 1);
      }
    });
  }

  if (!base.special.length) {
    base.special = ['特記すべき事項はありません。'];
  }

  return base;
}

function buildDoctorReportTemplate_(header, context, statusSections){
  const hospital = header?.hospital ? String(header.hospital).trim() : '';
  const doctor = header?.doctor ? String(header.doctor).trim() : '';
  const name = header?.name ? String(header.name).trim() : `ID:${header?.patientId || ''}`;
  const birth = header?.birth ? String(header.birth).trim() : '';
  const consent = context?.consentText ? String(context.consentText).trim() : '情報不足';
  const frequency = context?.frequencyLabel ? String(context.frequencyLabel).trim() : '情報不足';
  const rangeLabel = normalizeDoctorReportText_(context?.rangeLabel);
  const status = buildDoctorStatusFromSections_(statusSections);

  const body = ensureDoctorSentenceWithFallback_(
    status.body,
    rangeLabel
      ? `該当期間（${rangeLabel}）の記録では、心身機能の大きな変化は確認されていません。`
      : '心身機能の大きな変化は確認されていません。'
  );

  const activities = ensureDoctorSentenceWithFallback_(
    status.activities,
    '日常生活動作は概ね維持されています。'
  );

  const env = normalizeDoctorReportText_(status.environment);
  let participationSource = normalizeDoctorReportText_(status.participation);
  if (env) {
    participationSource = [participationSource, `環境・支援：${env}`].filter(Boolean).join(' / ');
  }
  const participation = ensureDoctorSentenceWithFallback_(
    participationSource,
    '社会参加や外出状況に大きな変化はありません。'
  );

  let safetySource = normalizeDoctorReportText_(status.safety);
  let safety = ensureDoctorSentenceWithFallback_(
    safetySource,
    '重大なリスクはみられず、訪問ごとにバイタルを確認しています。'
  );
  const complianceSentence = '同意内容に沿った施術を継続しております。';
  if (safety.indexOf(complianceSentence) < 0) {
    const trimmed = safety.replace(/[。．]+$/, '');
    safety = trimmed ? `${trimmed}。${complianceSentence}` : complianceSentence;
  }

  const specialList = normalizeDoctorSpecialList_(status.special).slice(0, 3);
  const special = (specialList
    .map(item => {
      const sentence = ensureDoctorSentenceWithFallback_(item, '');
      if (!sentence) return '';
      return `・${sentence}`;
    })
    .filter(Boolean)
    .join('\n')) || '・特記すべき事項はありません。';

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const createdAt = Utilities.formatDate(new Date(), tz, 'yyyy年M月d日');

return [
  `【病院名】${hospital || '不明'}`,
  `【担当医名】${doctor || '不明'}`,
  `【患者氏名】${name || '—'}`,
  `【生年月日】${birth || '不明'}`,
  `【同意内容】${consent || '情報不足'}`,
  `【施術頻度】${frequency || '情報不足'}`,
  '',
  '【患者の状態・経過】',
  // AI生成部分：痛みの状態、比較対象、ADL変化、新たな訴え、方針
  body
    ? body
    : '（情報不足のため生成できません）',
  '',
  '【特記すべき事項】',
  // AI抽出部分：リスク・体調管理＋末尾に必ず同意内容に沿った施術を継続しております。」
  (safety && !safety.includes('同意内容に沿った施術を継続しております。'))
    ? `${safety} 同意内容に沿った施術を継続しております。`
    : (safety || '特記すべき事項はありません。 同意内容に沿った施術を継続しております。'),
  '',
  `作成日：${createdAt}`,
  'べるつりー鍼灸マッサージ院',
  '東京都八王子市下柚木３－７－２－４０１',
  '042-682-2839',
  'mail:belltree@belltree1102.com'
].join('\n');
}

function buildHandoverDigestForSummary_(handovers, audience){
  if (!Array.isArray(handovers) || !handovers.length) return '';
  const latest = handovers.filter(h => String(h?.note || '').trim()).slice(-3);
  if (!latest.length) return '';
  const entries = latest.map(h => `${h.when || ''} ${String(h.note).trim()}`.trim());
  if (!entries.length) return '';
  const joined = entries.join(' / ');
  if (audience === 'doctor') {
    return `最近の申し送りでは、${joined}。`;
  }
  if (audience === 'caremanager') {
    return `申し送りの要点：${joined}。`;
  }
  return `最近のようす：${joined}。`;
}

function resolveReportTypeMeta_(reportType){
  const normalized = String(reportType || 'doctor').trim();
  const key = normalized.toLowerCase();
  switch (key) {
    case 'doctor':
      return { key: 'doctor', label: '医師向け報告書', specialLabel: '特記すべき事項' };
    case 'caremanager':
    case 'care_manager':
    case 'care-manager':
      return { key: 'caremanager', label: 'ケアマネ向けサマリ', specialLabel: '' };
    case 'family':
      return { key: 'family', label: '家族向けサマリ', specialLabel: '' };
    default:
      return { key, label: 'サマリ', specialLabel: '' };
  }
}

function normalizeAudienceRange_(rangeInput){
  const raw = String(rangeInput || '').trim();
  if (!raw) return 'all';
  const lower = raw.toLowerCase();
  const map = {
    '直近1か月': '1m',
    '直近１か月': '1m',
    '直近1ヶ月': '1m',
    '直近１ヶ月': '1m',
    '1m': '1m',
    'one_month': '1m',
    '直近2か月': '2m',
    '直近２か月': '2m',
    '直近2ヶ月': '2m',
    '直近２ヶ月': '2m',
    '2m': '2m',
    'two_month': '2m',
    '直近3か月': '3m',
    '直近３か月': '3m',
    '直近3ヶ月': '3m',
    '直近３ヶ月': '3m',
    '3m': '3m',
    'three_month': '3m',
    '全期間': 'all',
    'all': 'all',
    '直近6か月': '6m',
    '直近６か月': '6m',
    '直近6ヶ月': '6m',
    '直近６ヶ月': '6m',
    '6m': '6m',
    'six_month': '6m',
    '直近12か月': '12m',
    '直近１２か月': '12m',
    '直近12ヶ月': '12m',
    '直近１２ヶ月': '12m',
    '12m': '12m',
    'twelve_month': '12m',
    'custom': 'custom',
    'カスタム': 'custom'
  };
  if (map[raw]) return map[raw];
  if (map[lower]) return map[lower];
  const match = raw.match(/直近\s*(\d+)\s*(?:か月|ヶ月|か?\s*月)/);
  if (match) {
    const months = Math.max(1, Number(match[1] || 1));
    return `${months}m`;
  }
  return raw;
}

function normalizeRangeInputObject_(rangeInput){
  if (rangeInput == null || rangeInput === '') {
    return { key: 'all' };
  }
  if (typeof rangeInput === 'object') {
    const keyCandidate = rangeInput.key != null ? rangeInput.key
      : rangeInput.range != null ? rangeInput.range
        : rangeInput.value != null ? rangeInput.value
          : rangeInput.label;
    const key = normalizeAudienceRange_(keyCandidate);
    const normalized = { key: key || 'all' };
    if (normalized.key === 'custom') {
      normalized.start = rangeInput.start || rangeInput.startDate || rangeInput.from || '';
      normalized.end = rangeInput.end || rangeInput.endDate || rangeInput.to || '';
    }
    if (rangeInput.months != null) {
      normalized.months = rangeInput.months;
    }
    return normalized;
  }

  const raw = String(rangeInput || '').trim();
  if (!raw) {
    return { key: 'all' };
  }
  const key = normalizeAudienceRange_(raw);
  if (key === 'custom') {
    const normalized = { key: 'custom', start: '', end: '' };
    const customPrefixPattern = /^(custom|カスタム)[\s:：-]*/i;
    const rest = raw.replace(customPrefixPattern, '').trim();
    if (rest) {
      const tokens = rest.split(/[~〜\-–—|,、\s]+/).map(t => t.trim()).filter(Boolean);
      if (tokens.length >= 1) normalized.start = tokens[0];
      if (tokens.length >= 2) normalized.end = tokens[1];
    }
    return normalized;
  }
  return { key: key || 'all' };
}

function buildAiReportPrompt_(header, context){
  const lines = [];
  const rangeLabel = context?.range?.label || '全期間';
  lines.push('【患者情報】');
  lines.push(`- 氏名: ${header?.name || `ID:${header?.patientId || ''}`}`);
  lines.push(`- 施術録番号: ${header?.patientId || ''}`);
  if (header?.birth) lines.push(`- 生年月日: ${header.birth}`);
  if (header?.hospital) lines.push(`- 主治医/医療機関: ${header.hospital}${header?.doctor ? ` ${header.doctor}` : ''}`);
  if (header?.share) lines.push(`- 負担割合: ${header.share}`);
  lines.push(`- 対象期間: ${rangeLabel}`);

  const sections = Array.isArray(context?.sections) ? context.sections : [];
  if (sections.length) {
    lines.push('【AI下書きセクション】');
    sections.forEach(section => {
      const label = String(section?.label || section?.key || '').trim();
      const text = String(section?.text || '').trim();
      if (!label || !text) return;
      lines.push(`- ${label}: ${text}`);
    });
  }

  const notes = Array.isArray(context?.notes) ? context.notes : [];
  if (notes.length) {
    lines.push('【施術録メモ（古い順に最大12件）】');
    notes.slice(-12).forEach(note => {
      const when = String(note?.when || '').trim();
      const body = String(note?.note || note?.raw || '').trim();
      const vitals = String(note?.vitals || '').trim();
      const summary = [body, vitals ? `Vitals: ${vitals}` : '']
        .filter(Boolean)
        .join(' / ');
      lines.push(`- ${when}: ${summary}`);
    });
  }

  const handovers = Array.isArray(context?.handovers) ? context.handovers : [];
  if (handovers.length) {
    lines.push('【申し送り（古い順に最大10件）】');
    handovers.slice(-10).forEach(entry => {
      const when = String(entry?.when || '').trim();
      const note = String(entry?.note || '').trim();
      lines.push(`- ${when}: ${note}`);
    });
  }

  return lines.join('\n');
}

function buildAiReportSystemPrompt_(reportType){
  switch (String(reportType || '').toLowerCase()) {
    case 'doctor':
      return 'あなたは訪問マッサージ事業所のスタッフとして、主治医へ提出する訪問報告書を日本語で作成します。専門的で簡潔な医療文書とし、同意内容や施術頻度に触れつつ、心身機能・活動・社会参加・環境・リスクを整理してください。JSONで応答し、textとspecial(任意の配列)のみを含めます。';
    case 'caremanager':
      return 'あなたは訪問マッサージ事業所のスタッフとして、ケアマネジャー向けの報告サマリを日本語で作成します。介護支援専門員がサービス調整に使えるよう、状態変化と支援提案をわかりやすくまとめてください。JSONで応答し、textのみを含めます。';
    case 'family':
      return 'あなたは訪問マッサージ事業所のスタッフとして、ご家族向けのやさしい口調の報告文を日本語で作成します。安心感を与えつつ、様子と注意点を簡潔に伝えてください。JSONで応答し、textのみを含めます。';
    default:
      return 'あなたは訪問マッサージ事業所のスタッフとして、用途に合わせた報告文を日本語で作成します。JSONで応答し、textのみを含めます。';
  }
}

function generateAiSummaryServer(patientId, rangeKey, audience) {
  const range = resolveIcfSummaryRange_(rangeKey);
  const source = buildIcfSource_(patientId, range);
  const audienceMeta = resolveAudienceMeta_(audience);

  if (!source.patientFound) {
    return {
      ok: false,
      usedAi: true,
      audience: audienceMeta.key,
      audienceLabel: audienceMeta.label,
      text: '患者が見つかりませんでした。',
      meta: { patientFound: false, rangeLabel: range.label }
    };
  }

  const header = Object.assign({ patientId }, source.header || {});
  const context = {
    consentText: source.consent,
    frequencyLabel: source.frequencyLabel,
    rangeLabel: range.label,
    range,
    notes: source.notes,
    handovers: source.handovers
  };

  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const formatDate = (date) => {
    if (!(date instanceof Date) || isNaN(date.getTime())) return '';
    return Utilities.formatDate(date, timezone, 'yyyy-MM-dd');
  };

  let referenceReport = null;
  if (audienceMeta.key === 'doctor') {
    const latestHandover = getLatestHandoverEntry_(patientId);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    if (!isRecentHandoverEntry_(latestHandover, today)) {
      return {
        ok: false,
        usedAi: false,
        audience: audienceMeta.key,
        audienceLabel: audienceMeta.label,
        text: '申し送りが未入力のため、報告書を生成できません。申し送りを入力してください。',
        meta: {
          patientFound: true,
          rangeLabel: range.label,
          handoverRequired: true
        }
      };
    }
    referenceReport = findLatestDoctorReportEntry_(header.patientId);
    if (referenceReport && referenceReport.text) {
      context.previousDoctorReport = {
        text: referenceReport.text,
        when: referenceReport.when,
        ts: referenceReport.ts,
        rangeLabel: referenceReport.rangeLabel,
        rowNumber: referenceReport.rowNumber
      };
    }
  }

  const aiRes = composeAiReportViaOpenAI_(header, context, audienceMeta.key) || {};
  const text = typeof aiRes === 'object' ? (aiRes.text || '') : String(aiRes || '');
  const usedAi = !(aiRes && aiRes.via === 'local');

  const baseMeta = {
    patientFound: true,
    rangeLabel: range.label,
    rangeKey: range.key,
    rangeMonths: range.months,
    rangeStart: formatDate(range.startDate),
    rangeEnd: formatDate(range.endDate),
    noteCount: Array.isArray(source.notes) ? source.notes.length : 0,
    handoverCount: Array.isArray(source.handovers) ? source.handovers.length : 0,
    generationMode: usedAi ? 'AI' : 'ローカル整形'
  };
  if (referenceReport && referenceReport.rowNumber != null) {
    baseMeta.referenceReportId = String(referenceReport.rowNumber);
  }

  const result = {
    ok: true,
    usedAi,
    audience: audienceMeta.key,
    audienceLabel: audienceMeta.label,
    text,
    meta: baseMeta
  };

  if (aiRes && typeof aiRes === 'object' && aiRes.special != null) {
    result.special = aiRes.special;
  }

  const saved = persistAiReportsBatch_(header.patientId, range.label, [result]);
  if (saved && saved.length) {
    result.savedAt = saved[0].ts;
    result.persisted = true;
  }

  return result;
}
/***** OpenAI で AI レポート生成 *****/
function composeAiReportViaOpenAI_(header, context, audienceKey) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY が設定されていません。');
  }

  const promptConfig = buildReportPrompt_(header, context, audienceKey);
  const promptObject = typeof promptConfig === 'string' ? { userPrompt: promptConfig } : (promptConfig || {});
  const systemPrompt = promptObject.systemPrompt || SystemPrompt_GenericReport_JP;
  const userPrompt = promptObject.userPrompt || promptObject.prompt || '';
  if (!userPrompt) {
    throw new Error('AIプロンプトの生成に失敗しました。');
  }

  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: APP.OPENAI_MODEL || 'gpt-4o-mini', // または gpt-4o / gpt-4.1 など
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    temperature: 0.4
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const data = JSON.parse(res.getContentText());

  const text = data.choices?.[0]?.message?.content?.trim() || '';
  return { text, via: 'openai' };
}

/***** AI に渡すプロンプトを組み立てる *****/
function buildReportPrompt_(header, context, audienceKey) {
  const safe = (value) => {
    const text = value == null || value === '' ? '—' : String(value);
    return text.trim() ? text : '—';
  };

  if (audienceKey === 'doctor') {
    const formatEntries = (items, options) => {
      const opts = options || {};
      if (!Array.isArray(items) || !items.length) {
        return 'なし';
      }
      return items
        .slice(-10)
        .reverse()
        .map(entry => {
          const when = entry && entry.when ? `[${entry.when}]` : '';
          const pieces = [];
          const noteText = entry && typeof entry.note === 'string' && entry.note.trim()
            ? entry.note.trim()
            : (entry && typeof entry.raw === 'string' && entry.raw.trim() ? entry.raw.trim() : '');
          if (noteText) pieces.push(noteText);
          if (opts.includeVitals && entry && entry.vitals) {
            pieces.push(`バイタル: ${String(entry.vitals).trim()}`);
          }
          const body = pieces.filter(Boolean).join(' ／ ');
          const text = [when, body].filter(Boolean).join(' ');
          return `- ${text}`.trim();
        })
        .join('\n');
    };

    const lines = [];
    lines.push(`【医療機関】${safe(header && header.hospital)}`);
    lines.push(`【担当医】${safe(header && header.doctor)}`);
    lines.push(`【患者氏名】${safe(header && header.name)}`);
    lines.push(`【生年月日】${safe(header && header.birth)}`);
    lines.push(`【対象期間】${safe(context && context.rangeLabel)}`);
    lines.push(`【同意内容】${safe(context && context.consentText)}`);
    lines.push(`【施術頻度】${safe(context && context.frequencyLabel)}`);
    lines.push('');
    lines.push('【申し送り（最新順）】');
    lines.push(formatEntries(context && context.handovers, { includeVitals: false }));
    lines.push('');
    lines.push('【施術録メモ（最新順）】');
    lines.push(formatEntries(context && context.notes, { includeVitals: true }));
    lines.push('');
    lines.push('上記情報をもとに医師向け施術報告書を作成してください。');
    lines.push('必要に応じてVASやADLなどの客観指標を強調して構成してください。');
    if (context && context.previousDoctorReport && context.previousDoctorReport.text) {
      const previous = context.previousDoctorReport;
      const headerParts = [];
      if (previous.rangeLabel) headerParts.push(`対象期間：${String(previous.rangeLabel).trim()}`);
      if (previous.when) headerParts.push(`作成日時：${String(previous.when).trim()}`);
      lines.push('');
      lines.push(headerParts.length ? `【前回報告書】${headerParts.join(' ｜ ')}` : '【前回報告書】');
      lines.push('---');
      lines.push(String(previous.text).trim());
      lines.push('---');
      lines.push('前回内容を踏まえつつ、重複表現を避けて最新の経過を反映してください。');
    }
    return {
      systemPrompt: SystemPrompt_DoctorReport_JP,
      userPrompt: lines.join('\n')
    };
  }

  const roleLabel = audienceKey === 'doctor'
    ? '医師'
    : audienceKey === 'caremanager'
      ? 'ケアマネジャー'
      : 'ご家族';

  const defaultLines = [];
  defaultLines.push(`【病院名】${safe(header && header.hospital)}`);
  defaultLines.push(`【担当医名】${safe(header && header.doctor)}`);
  defaultLines.push(`【患者氏名】${safe(header && header.name)}`);
  defaultLines.push(`【生年月日】${safe(header && header.birth)}`);
  defaultLines.push(`【同意内容】${safe(context && context.consentText)}`);
  defaultLines.push(`【施術頻度】${safe(context && context.frequencyLabel)}`);
  defaultLines.push('');
  defaultLines.push(`${roleLabel}向けに患者様の状態・経過をまとめてください。`);
  defaultLines.push('必ず「同意内容に沿った施術を継続しております。」という一文を含めてください。');
  defaultLines.push('');
  defaultLines.push('参考情報：');
  defaultLines.push(`- Notes: ${JSON.stringify((context && context.notes) || [])}`);
  defaultLines.push(`- Handovers: ${JSON.stringify((context && context.handovers) || [])}`);
  defaultLines.push(`- 期間: ${safe(context && context.rangeLabel)}`);

  return {
    systemPrompt: SystemPrompt_GenericReport_JP,
    userPrompt: defaultLines.join('\n')
  };
}


function composeAiReportLocal_(header, context, reportType){
  const audienceMeta = resolveAudienceMeta_(reportType);
  const range = context?.range || { startDate: null, endDate: new Date(), label: '全期間' };
  const sections = Array.isArray(context?.sections) ? context.sections : [];
  const source = context?.source || { header, notes: [], handovers: [] };
  const text = buildAudienceNarrative_(audienceMeta, header, range, source, sections);
  let special = [];
  if (audienceMeta.key === 'doctor') {
    special = normalizeDoctorSpecialList_(buildDoctorStatusFromSections_(sections).special || []);
  }
  return { via: 'local', audience: audienceMeta.key, text, special };
}

function normalizeReportSpecial_(special){
  if (special == null) return [];
  if (Array.isArray(special)) {
    return special
      .map(item => normalizeDoctorReportText_(item))
      .filter(Boolean);
  }
  if (typeof special === 'object') {
    if (Array.isArray(special.special)) {
      return normalizeReportSpecial_(special.special);
    }
    return [];
  }
  return String(special || '')
    .split(/\r?\n|[,、・]/)
    .map(item => normalizeDoctorReportText_(item))
    .filter(Boolean);
}

function parseReportSpecialText_(value){
  return normalizeReportSpecial_(value);
}

function parseReportStatusMeta_(status){
  const meta = {
    usedAi: null,
    noteCount: null,
    handoverCount: null
  };
  const text = String(status || '').trim();
  if (!text) return meta;
  text.split('|')
    .map(part => part.trim())
    .filter(Boolean)
    .forEach(part => {
      const [rawKey, rawValue] = part.split('=');
      const key = (rawKey || '').trim().toLowerCase();
      const value = (rawValue || '').trim();
      if (!key) return;
      if (key === 'via') {
        meta.usedAi = value !== 'local';
        return;
      }
      const num = Number(value);
      if (Number.isFinite(num)) {
        if (key === 'notes') meta.noteCount = num;
        if (key === 'handovers') meta.handoverCount = num;
      }
    });
  return meta;
}

function resolveAudienceKeyFromAny_(keyCandidate, labelCandidate){
  const normalizedKey = String(keyCandidate || '').trim();
  if (normalizedKey) {
    const meta = resolveAudienceMeta_(normalizedKey);
    if (meta && meta.key) {
      return meta.key;
    }
  }
  const label = String(labelCandidate || '').trim();
  if (!label) {
    return normalizedKey.toLowerCase();
  }
  if (label === '医師向け報告書') return 'doctor';
  if (label === 'ケアマネ向けサマリ') return 'caremanager';
  if (label === '家族向けサマリ') return 'family';
  return normalizedKey.toLowerCase();
}

function persistAiReportsBatch_(patientId, rangeLabel, summaries){
  const normalized = normId_(patientId);
  if (!normalized || !Array.isArray(summaries) || !summaries.length) {
    return [];
  }

  const sheet = ensureAiReportSheet_();
  const rows = [];
  const saved = [];
  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const label = String(rangeLabel || '');

  summaries.forEach(summary => {
    if (!summary || summary.ok === false) return;
    const audienceMeta = resolveAudienceMeta_(summary.audience || '');
    const meta = summary.meta ? Object.assign({}, summary.meta) : {};
    let text = summary.text != null ? String(summary.text) : '';
    let doctorSectionsMeta = null;
    let specialList = normalizeReportSpecial_(summary.special);

    if (audienceMeta.key === 'doctor') {
      const normalizedDoctor = normalizeDoctorReportTextForStorage_(text);
      if (normalizedDoctor && normalizedDoctor.text) {
        text = normalizedDoctor.text;
        summary.text = text;
        doctorSectionsMeta = {
          section1: normalizedDoctor.section1,
          section2: normalizedDoctor.section2,
          section3: normalizedDoctor.section3,
          frequencyText: normalizedDoctor.frequencyText
        };
        meta.doctorSections = doctorSectionsMeta;
        if (!specialList.length) {
          specialList = normalizeDoctorSpecialList_(normalizedDoctor.section3);
          summary.special = specialList.slice();
        }
      }
    }

    summary.meta = meta;
    specialList = normalizeReportSpecial_(summary.special);
    const statusParts = [];
    statusParts.push(summary.usedAi === false ? 'via=local' : 'via=ai');
    if (meta.noteCount != null) statusParts.push(`notes=${meta.noteCount}`);
    if (meta.handoverCount != null) statusParts.push(`handovers=${meta.handoverCount}`);
    const status = statusParts.join(' | ');
    const specialText = specialList.join('\n');
    const ts = new Date();
    const rangeText = label || String(meta.rangeLabel || '');
    let periodValue = '';
    if (meta.rangeMonths != null && meta.rangeMonths !== '') {
      periodValue = String(meta.rangeMonths);
    } else if (meta.periodMonths != null && meta.periodMonths !== '') {
      periodValue = String(meta.periodMonths);
    } else if (meta.rangeKey) {
      const keyText = String(meta.rangeKey);
      if (/^\d+m$/.test(keyText)) {
        periodValue = keyText.replace('m', '');
      } else if (keyText === 'all') {
        periodValue = 'all';
      } else if (keyText === 'custom') {
        periodValue = 'custom';
      }
    }
    const referenceReportId = meta.referenceReportId != null ? String(meta.referenceReportId) : '';
    const generationMode = meta.generationMode
      ? String(meta.generationMode)
      : (summary.usedAi === false ? 'ローカル整形' : 'AI');
    meta.rangeLabel = rangeText;
    meta.rangeKey = meta.rangeKey || (periodValue && periodValue !== 'all' && periodValue !== 'custom' ? `${periodValue}m` : (periodValue || ''));
    meta.rangeMonths = periodValue;
    meta.referenceReportId = referenceReportId;
    meta.generationMode = generationMode;
    rows.push([
      ts,
      String(normalized),
      rangeText,
      audienceMeta.label,
      audienceMeta.key,
      text,
      status,
      specialText,
      periodValue,
      referenceReportId,
      generationMode
    ]);
    const savedMeta = Object.assign({}, meta, {
      rangeLabel: rangeText,
      noteCount: meta.noteCount != null ? Number(meta.noteCount) : null,
      handoverCount: meta.handoverCount != null ? Number(meta.handoverCount) : null
    });
    saved.push({
      ts: ts.getTime(),
      when: Utilities.formatDate(ts, timezone, 'yyyy-MM-dd HH:mm'),
      rangeLabel: rangeText,
      audience: audienceMeta.key,
      audienceLabel: audienceMeta.label,
      text,
      status,
      special: specialList,
      usedAi: summary.usedAi === false ? false : true,
      meta: savedMeta
    });
  });

  if (!rows.length) {
    return [];
  }

  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, AI_REPORT_SHEET_HEADER.length).setValues(rows);
  invalidatePatientCaches_(normalized, { reports: true });
  return saved;
}

function fetchReportHistoryForPid_(normalized){
  if (!normalized) return [];
  const sheet = ensureAiReportSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const width = Math.max(sheet.getLastColumn(), AI_REPORT_SHEET_HEADER.length);
  const range = sheet.getRange(2, 1, lastRow - 1, width);
  const values = range.getValues();
  const displays = range.getDisplayValues();
  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const disp = displays[i];
    const sheetRow = 2 + i;
    const pidRaw = row[1] != null && row[1] !== '' ? row[1] : disp[1];
    const pid = normId_(pidRaw);
    if (pid !== normalized) continue;
    const tsRaw = row[0];
    let ts = 0;
    if (tsRaw instanceof Date) {
      ts = tsRaw.getTime();
    } else if (typeof tsRaw === 'number') {
      ts = tsRaw;
    } else if (tsRaw) {
      const parsed = new Date(tsRaw);
      if (!isNaN(parsed.getTime())) ts = parsed.getTime();
    }
    const whenText = disp[0] || (ts ? Utilities.formatDate(new Date(ts), timezone, 'yyyy-MM-dd HH:mm') : '');
    const rangeLabel = disp[2] || row[2] || '';
    const audienceLabel = disp[3] || row[3] || '';
    const audienceKey = resolveAudienceKeyFromAny_(row[4] || '', audienceLabel);
    const text = row[5] != null ? String(row[5]) : (disp[5] || '');
    const status = row[6] != null ? String(row[6]) : (disp[6] || '');
    const special = parseReportSpecialText_(row[7] != null ? row[7] : disp[7]);
    const periodRaw = row.length > 8 && row[8] != null ? row[8] : (disp.length > 8 ? disp[8] : '');
    const periodMonths = periodRaw != null && periodRaw !== '' ? String(periodRaw) : '';
    const referenceRaw = row.length > 9 && row[9] != null ? row[9] : (disp.length > 9 ? disp[9] : '');
    const referenceReportId = referenceRaw != null && referenceRaw !== '' ? String(referenceRaw) : '';
    const modeRaw = row.length > 10 && row[10] != null ? row[10] : (disp.length > 10 ? disp[10] : '');
    const parsedStatus = parseReportStatusMeta_(status);
    const generationMode = modeRaw != null && modeRaw !== ''
      ? String(modeRaw)
      : (parsedStatus.usedAi === false ? 'ローカル整形' : 'AI');
    const derivedRangeKey = periodMonths
      ? (periodMonths === 'all'
        ? 'all'
        : periodMonths === 'custom'
          ? 'custom'
          : `${periodMonths}m`)
      : '';
    rows.push({
      rowNumber: sheetRow,
      ts,
      when: whenText,
      rangeLabel,
      audience: audienceKey,
      audienceLabel: audienceLabel || getIcfAudienceLabel_(audienceKey),
      text,
      status,
      special,
      usedAi: parsedStatus.usedAi == null ? true : !!parsedStatus.usedAi,
      meta: {
        rangeLabel,
        noteCount: parsedStatus.noteCount,
        handoverCount: parsedStatus.handoverCount,
        rangeMonths: periodMonths,
        rangeKey: derivedRangeKey,
        referenceReportId,
        generationMode: generationMode || (parsedStatus.usedAi === false ? 'ローカル整形' : 'AI')
      }
    });
  }
  return rows.sort((a, b) => (b.ts || 0) - (a.ts || 0));
}

function findLatestDoctorReportEntry_(patientId){
  const normalized = normId_(patientId);
  if (!normalized) return null;
  const history = fetchReportHistoryForPid_(normalized);
  if (!Array.isArray(history) || !history.length) return null;
  for (let i = 0; i < history.length; i++) {
    const entry = history[i];
    if (entry && entry.audience === 'doctor' && entry.text && String(entry.text).trim()) {
      return entry;
    }
  }
  return null;
}

function listPatientReports(patientId) {
  const normalized = normId_(patientId);
  if (!normalized) {
    return { ok: false, message: '患者IDが指定されていません。', reports: [] };
  }
  const reports = cacheFetch_(PATIENT_CACHE_KEYS.reports(normalized), () => fetchReportHistoryForPid_(normalized), PATIENT_CACHE_TTL_SECONDS) || [];
  return { ok: true, patientId: normalized, reports };
}

function getSavedReportsForUI(patientId) {
  const normalized = normId_(patientId);
  if (!normalized) {
    return { ok: false, message: '患者IDが指定されていません。', reports: {} };
  }
  const history = cacheFetch_(PATIENT_CACHE_KEYS.reports(normalized), () => fetchReportHistoryForPid_(normalized), PATIENT_CACHE_TTL_SECONDS) || [];
  const latestByAudience = {};
  history.forEach(entry => {
    if (!entry || !entry.audience) return;
    const current = latestByAudience[entry.audience];
    if (!current || (entry.ts || 0) > (current.ts || 0)) {
      latestByAudience[entry.audience] = entry;
    }
  });
  const reports = {};
  let latestTs = 0;
  Object.keys(latestByAudience).forEach(key => {
    const entry = latestByAudience[key];
    reports[key] = {
      text: entry.text || '',
      audience: entry.audience,
      audienceLabel: entry.audienceLabel,
      when: entry.when,
      ts: entry.ts,
      rangeLabel: entry.rangeLabel,
      meta: entry.meta,
      usedAi: entry.usedAi,
      special: entry.special
    };
    if ((entry.ts || 0) > latestTs) {
      latestTs = entry.ts || 0;
    }
  });
  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const latestWhen = latestTs ? Utilities.formatDate(new Date(latestTs), timezone, 'yyyy-MM-dd HH:mm') : '';
  const representative = Object.values(reports)[0];
  return {
    ok: true,
    patientId: normalized,
    reports,
    rangeLabel: representative ? (representative.rangeLabel || '') : '',
    latestWhen
  };
}

function updateAiReportEntry(payload) {
  const rowNumber = Number(payload && payload.rowNumber);
  if (!rowNumber || rowNumber < 2) {
    throw new Error('rowNumberが不正です');
  }
  const sheet = ensureAiReportSheet_();
  const lastRow = sheet.getLastRow();
  if (rowNumber > lastRow) {
    throw new Error('指定された行が存在しません');
  }
  const width = AI_REPORT_SHEET_HEADER.length;
  const values = sheet.getRange(rowNumber, 1, 1, width).getValues()[0];
  const pid = normId_(values[1]);
  if (!pid) {
    throw new Error('患者IDを特定できません');
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const now = new Date();
  const text = payload && payload.text != null ? String(payload.text) : '';
  const rangeLabel = payload && payload.rangeLabel != null ? String(payload.rangeLabel) : null;

  sheet.getRange(rowNumber, 1).setValue(now);
  sheet.getRange(rowNumber, 6).setValue(text);
  if (rangeLabel != null) {
    sheet.getRange(rowNumber, 3).setValue(rangeLabel);
  }

  const statusRange = sheet.getRange(rowNumber, 7);
  const statusRaw = String(statusRange.getValue() || '');
  if (statusRaw.indexOf('edited=manual') < 0) {
    const updatedStatus = statusRaw ? `${statusRaw} | edited=manual` : 'edited=manual';
    statusRange.setValue(updatedStatus);
  }

  const generationRange = sheet.getRange(rowNumber, 11);
  try {
    generationRange.setValue('編集反映');
  } catch (err) {
    Logger.log('[updateAiReportEntry] failed to set generation mode: ' + (err && err.message ? err.message : err));
  }

  invalidatePatientCaches_(pid, { reports: true });

  return {
    ok: true,
    patientId: pid,
    rowNumber,
    when: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm'),
    text,
    rangeLabel: rangeLabel != null ? rangeLabel : values[2]
  };
}

function duplicateAiReportEntry(payload) {
  const rowNumber = Number(payload && payload.rowNumber);
  if (!rowNumber || rowNumber < 2) {
    throw new Error('rowNumberが不正です');
  }
  const sheet = ensureAiReportSheet_();
  const lastRow = sheet.getLastRow();
  if (rowNumber > lastRow) {
    throw new Error('指定された行が存在しません');
  }
  const width = AI_REPORT_SHEET_HEADER.length;
  const values = sheet.getRange(rowNumber, 1, 1, width).getValues()[0];
  const pid = normId_(values[1]);
  if (!pid) {
    throw new Error('患者IDを特定できません');
  }

  const sourceRangeLabel = values[2] != null ? String(values[2]) : '';
  const sourceAudienceLabel = values[3] != null ? String(values[3]) : '';
  const sourceAudienceKey = values[4] != null ? String(values[4]) : '';
  const sourceText = values[5] != null ? String(values[5]) : '';
  const sourceStatus = values[6] != null ? String(values[6]) : '';
  const sourceSpecial = values[7] != null ? values[7] : '';

  const audienceInput = payload && (payload.audienceKey || payload.audience || payload.targetAudience);
  const audienceLabelInput = payload && payload.audienceLabel;
  const resolvedAudienceKey = audienceInput
    ? resolveAudienceKeyFromAny_(audienceInput, audienceLabelInput)
    : resolveAudienceKeyFromAny_(sourceAudienceKey, sourceAudienceLabel);
  const audienceMeta = resolveAudienceMeta_(resolvedAudienceKey);

  const rangeLabel = payload && payload.rangeLabel != null ? String(payload.rangeLabel) : sourceRangeLabel;
  const text = payload && payload.text != null ? String(payload.text) : sourceText;
  const statusBase = payload && payload.status ? String(payload.status) : sourceStatus;
  const status = statusBase ? `${statusBase} | copied=manual` : 'copied=manual';

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const now = new Date();

  const periodValue = values.length > 8 ? values[8] : '';
  const referenceId = values.length > 9 ? values[9] : '';
  const generationMode = '再生成';

  sheet.appendRow([
    now,
    pid,
    rangeLabel,
    audienceMeta.label,
    audienceMeta.key,
    text,
    status,
    sourceSpecial,
    periodValue,
    referenceId,
    generationMode
  ]);

  invalidatePatientCaches_(pid, { reports: true });

  const newRowNumber = sheet.getLastRow();

  return {
    ok: true,
    patientId: pid,
    rowNumber: newRowNumber,
    when: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm'),
    audience: audienceMeta.key,
    rangeLabel,
    text
  };
}

function clearDoctorReportReminder(payload) {
  const pid = String(payload && payload.patientId || '').trim();
  if (!pid) {
    throw new Error('patientIdが空です');
  }
  const newsType = String(payload && payload.newsType || '同意').trim() || '同意';
  const metaMatches = { type: 'consent_verification' };
  if (payload && payload.consentExpiry) {
    metaMatches.consentExpiry = String(payload.consentExpiry);
  }
  const options = {
    metaType: 'consent_verification',
    metaMatches
  };
  if (payload && payload.newsMessage) {
    options.messageContains = String(payload.newsMessage);
  }
  if (payload && typeof payload.newsRow === 'number') {
    options.rowNumber = Number(payload.newsRow);
  }
  const cleared = markNewsClearedByType(pid, newsType, options);
  return { ok: true, cleared };
}

function buildIcfSource_(pid, range){
  const header = getPatientHeader(pid);
  if (!header) {
    return { patientFound: false };
  }
  const consentText = (header && typeof header.consentContent === 'string')
    ? header.consentContent.trim()
    : getConsentContentForPatient_(pid);
  const effectiveEndDate = range && range.endDate instanceof Date ? range.endDate : new Date();
  const frequencyLabel = determineTreatmentFrequencyLabel_(
    countTreatmentsInRecentMonth_(pid, effectiveEndDate)
  );
  const notes = getTreatmentNotesInRange_(pid, range.startDate, range.endDate);
  const handovers = getHandoversInRange_(pid, range.startDate, range.endDate);
  return {
    patientFound: true,
    header,
    notes,
    handovers,
    consent: consentText || '',
    frequencyLabel
  };
}

function resolveAudienceMeta_(audience){
  const key = String(audience || '').toLowerCase();
  switch (key) {
    case 'doctor':
      return { key: 'doctor', label: '医師向け報告書' };
    case 'caremanager':
    case 'care_manager':
    case 'care-manager':
      return { key: 'caremanager', label: 'ケアマネ向けサマリ' };
    case 'family':
      return { key: 'family', label: '家族向けサマリ' };
    default:
      return { key, label: 'サマリ' };
  }
}

function summarizeSectionsForAudience_(audienceKey, sections){
  const texts = (Array.isArray(sections) ? sections : [])
    .map(sec => `${sec.label}：${sec.text}`)
    .filter(Boolean);
  if (!texts.length) return '';
  if (audienceKey === 'family') {
    return texts.join('\n');
  }
  return texts.join('\n');
}

function buildAudienceNarrative_(audienceMeta, header, range, source, sections){
  const audienceKey = audienceMeta.key;
  const rangeLabel = range.label || '全期間';
  const handovers = Array.isArray(source.handovers) ? source.handovers : [];
  const sectionSummary = summarizeSectionsForAudience_(audienceKey, sections);
  const handoverDigest = buildHandoverDigestForSummary_(handovers, audienceKey);

  if (audienceKey === 'doctor') {
    const context = {
      consentText: getConsentContentForPatient_(header.patientId),
      frequencyLabel: determineTreatmentFrequencyLabel_(countTreatmentsInRecentMonth_(header.patientId, range.endDate)),
      rangeLabel
    };
    return buildDoctorReportTemplate_(header, context, sections);
  }

  if (audienceKey === 'caremanager') {
    const lines = [];
    lines.push(`【対象期間】${rangeLabel}`);
    lines.push(`【ご利用者】${header.name || `ID:${header.patientId}`}`);
    if (sectionSummary) {
      lines.push('【状態と変化】');
      lines.push(sectionSummary);
    } else {
      lines.push('【状態と変化】該当期間の記録が少なく、明確な変化は確認できませんでした。');
    }
    if (handoverDigest) lines.push(handoverDigest);
    return lines.join('\n');
  }

  const lines = [];
  const displayName = header.name || 'ご利用者さま';
  lines.push(`${displayName}のご様子（${rangeLabel}）をご報告します。`);
  if (sectionSummary) {
    lines.push(sectionSummary);
  } else {
    lines.push('この期間の詳細な記録は少ないですが、引き続き安全に配慮しながら訪問を継続しています。');
  }
  if (handoverDigest) lines.push(handoverDigest);
  lines.push('ご不明な点があればいつでもご連絡ください。');
  return lines.join('\n');
}

/**
 * 3種類まとめて生成（doctor / caremanager / family）
 */
function generateAllAiSummariesServer(patientId, rangeKey) {
  const range = resolveIcfSummaryRange_(rangeKey);
  const source = buildIcfSource_(patientId, range);

  if (!source.patientFound) {
    return {
      ok: false,
      usedAi: true,
      reports: null,
      meta: { patientFound: false, rangeLabel: range.label }
    };
  }

  const header = Object.assign({ patientId }, source.header || {});

  const context = {
    consentText: source.consent,
    frequencyLabel: source.frequencyLabel,
    rangeLabel: range.label,
    range,
    notes: source.notes,
    handovers: source.handovers
  };

  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const formatDate = (date) => {
    if (!(date instanceof Date) || isNaN(date.getTime())) return '';
    return Utilities.formatDate(date, timezone, 'yyyy-MM-dd');
  };

  const baseMeta = {
    patientFound: true,
    rangeLabel: range.label,
    rangeKey: range.key,
    rangeMonths: range.months,
    rangeStart: formatDate(range.startDate),
    rangeEnd: formatDate(range.endDate),
    noteCount: Array.isArray(source.notes) ? source.notes.length : 0,
    handoverCount: Array.isArray(source.handovers) ? source.handovers.length : 0,
    generationMode: 'AI'
  };

  const doctorContext = Object.assign({}, context);
  const previousDoctorReport = findLatestDoctorReportEntry_(header.patientId);
  if (previousDoctorReport && previousDoctorReport.text) {
    doctorContext.previousDoctorReport = {
      text: previousDoctorReport.text,
      when: previousDoctorReport.when,
      ts: previousDoctorReport.ts,
      rangeLabel: previousDoctorReport.rangeLabel,
      rowNumber: previousDoctorReport.rowNumber
    };
  }

  const doctorRes = composeAiReportViaOpenAI_(header, doctorContext, 'doctor') || {};
  const caremanagerRes = composeAiReportViaOpenAI_(header, context, 'caremanager') || {};
  const familyRes = composeAiReportViaOpenAI_(header, context, 'family') || {};

  const doctorMeta = Object.assign({}, baseMeta, {
    generationMode: !(doctorRes && doctorRes.via === 'local') ? 'AI' : 'ローカル整形'
  });
  if (previousDoctorReport && previousDoctorReport.rowNumber != null) {
    doctorMeta.referenceReportId = String(previousDoctorReport.rowNumber);
  }
  const caremanagerMeta = Object.assign({}, baseMeta, {
    generationMode: !(caremanagerRes && caremanagerRes.via === 'local') ? 'AI' : 'ローカル整形'
  });
  const familyMeta = Object.assign({}, baseMeta, {
    generationMode: !(familyRes && familyRes.via === 'local') ? 'AI' : 'ローカル整形'
  });

  const reports = {
    doctor: {
      ok: true,
      usedAi: !(doctorRes && doctorRes.via === 'local'),
      audience: 'doctor',
      audienceLabel: getIcfAudienceLabel_('doctor'),
      text: typeof doctorRes === 'object' ? (doctorRes.text || '') : String(doctorRes || ''),
      special: typeof doctorRes === 'object' ? doctorRes.special : undefined,
      meta: doctorMeta
    },
    caremanager: {
      ok: true,
      usedAi: !(caremanagerRes && caremanagerRes.via === 'local'),
      audience: 'caremanager',
      audienceLabel: getIcfAudienceLabel_('caremanager'),
      text: typeof caremanagerRes === 'object' ? (caremanagerRes.text || '') : String(caremanagerRes || ''),
      special: typeof caremanagerRes === 'object' ? caremanagerRes.special : undefined,
      meta: caremanagerMeta
    },
    family: {
      ok: true,
      usedAi: !(familyRes && familyRes.via === 'local'),
      audience: 'family',
      audienceLabel: getIcfAudienceLabel_('family'),
      text: typeof familyRes === 'object' ? (familyRes.text || '') : String(familyRes || ''),
      special: typeof familyRes === 'object' ? familyRes.special : undefined,
      meta: familyMeta
    }
  };

  const saved = persistAiReportsBatch_(header.patientId, range.label, Object.values(reports));
  if (saved && saved.length) {
    const savedMap = {};
    saved.forEach(entry => { savedMap[entry.audience] = entry; });
    Object.keys(reports).forEach(key => {
      const entry = savedMap[key];
      if (entry) {
        reports[key].savedAt = entry.ts;
        reports[key].persisted = true;
      }
    });
  }

  return {
    ok: true,
    usedAi: true,
    reports,
    rangeLabel: range.label,
    meta: baseMeta
  };
}

/**
 * フロントUI向け：まとめて取得（ラベル付き）
 */
function getReportsForUI(patientId, rangeInput) {
  const reports = generateAllAiSummariesServer(patientId, rangeInput);
  return {
    ok: !!reports.ok,
    usedAi: true,
    rangeLabel: reports.rangeLabel || reports?.meta?.rangeLabel || '',
    doctor: reports?.reports?.doctor?.text || '',
    caremanager: reports?.reports?.caremanager?.text || '',
    family: reports?.reports?.family?.text || '',
    reports
  };
}

/**
 * 個別レポート生成（従来の payload 形式をサポート）
 */
function generateAiReport(payload) {
  const meta = payload && typeof payload === 'object'
    ? resolveReportTypeMeta_(payload.reportType)
    : resolveReportTypeMeta_('');

  const patientId = payload?.patientId || payload?.pid || payload?.id || '';
  if (!patientId) {
    return {
      ok: false,
      usedAi: true,
      reportType: meta.key,
      message: '患者IDが指定されていません。'
    };
  }

  const rangeInput = payload && payload.range != null
    ? payload.range
    : (payload && payload.rangeKey != null ? payload.rangeKey : 'all');
  return generateAiSummaryServer(patientId, rangeInput, meta.key);
}

/**
 * オーディエンスの表示ラベル
 */
function getIcfAudienceLabel_(audience) {
  switch (audience) {
    case 'doctor': return '医師向け報告書';
    case 'caremanager': return 'ケアマネ向けサマリ';
    case 'family': return '家族向けサマリ';
    default: return 'サマリ';
  }
}

function ensureIntakeScaffolding_() {
  const wb = ss();
  // Intake_Staging が無ければ最低限のヘッダで作る（intakeGetValuesMap_ が読む前提）
  if (!wb.getSheetByName('Intake_Staging')) {
    const sh = wb.insertSheet('Intake_Staging');
    sh.getRange(1,1,1,9).setValues([[
      'leadId','ts','code','json','createdAt','updatedAt','author','mode','snapshot'
    ]]);
  }
  // LeadStatus はあなたの ensureIntakeSheets_() が面倒を見ているので触らない
}

/***** ── 差し替え：doGet ──*****/
function doGet(e) {
  e = e || {};

  if (shouldHandleDashboardApi_(e)) {
    const data = getDashboardData();
    return createJsonResponse_(data);
  }

  const path = (e && e.pathInfo ? String(e.pathInfo) : '').replace(/^\/+|\/+$/g, '').toLowerCase();
  let view = e.parameter ? (e.parameter.view || 'welcome') : 'welcome';
  if (path === 'treatmentapp') {
    view = 'record';
  }
  let templateFile = '';

  switch(view){
    case 'intake':       templateFile = 'intake'; break;
    case 'visit':        templateFile = 'intake'; break;
    case 'intake_list':  templateFile = 'intake_list'; break;
    case 'admin':        templateFile = 'admin'; break;
    case 'attendance':   templateFile = 'attendance'; break;
    case 'vacancy':      templateFile = 'vacancy'; break;
    case 'albyte':       templateFile = 'albyte'; break;
    case 'albyte_admin': templateFile = 'albyte_admin'; break;
    case 'albyte_report':templateFile = 'albyte_report'; break;
    case 'payroll':      templateFile = 'payroll'; break;
    case 'payroll_pdf_family': templateFile = 'payroll_pdf_family'; break;
    case 'billing':      templateFile = 'billing'; break;
    case 'dashboard':    templateFile = 'dashboard'; break;
    case 'record':       templateFile = 'app'; break;   // ★ app.html を record として表示
    case 'report':       templateFile = 'report'; break;
    default:             templateFile = 'welcome'; break;
  }

  const t = HtmlService.createTemplateFromFile(templateFile);

  // ここでURLを渡す
  t.baseUrl = ScriptApp.getService().getUrl();

  // 患者ID（?patientId=XXXX / ?id=XXXX）をテンプレートに渡す
  if (e.parameter && (e.parameter.patientId || e.parameter.id)) {
    t.patientId = e.parameter.patientId || e.parameter.id;
  } else {
    t.patientId = "";
  }
  t.payrollPdfData = {};

  if(e.parameter && e.parameter.lead) t.lead = e.parameter.lead;

  return t.evaluate()
           .setTitle('受付アプリ')
           .addMetaTag('viewport','width=device-width, initial-scale=1.0');
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

function notifyChat_(message){
  const url = (PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL') || '').trim();
  if (!url) { Logger.log('CHAT_WEBHOOK_URL 未設定'); return; }
  const payload = JSON.stringify({ text: message });
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: payload
  });
}

function decodeWebhookEmailKey_(key){
  const raw = String(key || '').trim();
  if (!raw) return '';
  if (looksLikeEmail_(raw)) {
    return raw.toLowerCase();
  }
  const upper = raw.toUpperCase();
  const patterns = [
    'CHAT_WEBHOOK_URL__',
    'CHAT_WEBHOOK_URL_',
    'CHAT_WEBHOOK__',
    'CHAT_WEBHOOK_',
    'WEBHOOK_URL__',
    'WEBHOOK_URL_',
    'WEBHOOK__',
    'WEBHOOK_'
  ];
  for (let i = 0; i < patterns.length; i++) {
    const prefix = patterns[i];
    if (upper.startsWith(prefix)) {
      const tail = raw.substring(prefix.length);
      const decoded = tail
        .replace(/(__AT__|_AT_|-AT-)/gi, '@')
        .replace(/(__DOT__|_DOT_|-DOT-)/gi, '.')
        .trim();
      if (decoded.indexOf('@') >= 0) {
        return decoded.toLowerCase();
      }
      if (tail.indexOf('@') >= 0) {
        return tail.toLowerCase();
      }
    }
  }
  return '';
}

function looksLikeEmail_(text){
  return /@/.test(text || '') && /\./.test(text || '');
}

function normalizeEmailKey_(email){
  return String(email || '').trim().toLowerCase();
}

function postJsonWebhook_(webhookUrl, payload){
  if (!webhookUrl) {
    throw new Error('Webhook URL が設定されていません');
  }
  const response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload || {})
  });
  const status = typeof response.getResponseCode === 'function' ? response.getResponseCode() : null;
  const body = typeof response.getContentText === 'function' ? response.getContentText() : '';
  if (status && status >= 400) {
    throw new Error(`Webhook status ${status}: ${body}`);
  }
  return { status, body };
}

function getWebhookConfig_(){
  const props = PropertiesService.getScriptProperties().getProperties() || {};
  const map = new Map();
  const defaultUrl = String((props.CHAT_WEBHOOK_URL_DEFAULT || props.CHAT_WEBHOOK_URL || '')).trim();
  Object.keys(props).forEach(key => {
    const value = String(props[key] || '').trim();
    if (!value) return;
    if (key === 'CHAT_WEBHOOK_URL_DEFAULT' || key === 'CHAT_WEBHOOK_URL') return;

    let email = '';
    if (looksLikeEmail_(key)) {
      email = key.toLowerCase();
    } else {
      email = decodeWebhookEmailKey_(key);
    }
    if (email) {
      map.set(email, value);
    }
  });
  return { map, defaultUrl };
}

function createStaffShiftRule_(identifier, options){
  const opts = options || {};
  const aliases = [];
  const normalizedId = normalizeEmailKey_(identifier);
  if (normalizedId) {
    aliases.push(normalizedId);
  }
  if (Array.isArray(opts.aliases)) {
    opts.aliases.forEach(value => {
      const alias = normalizeEmailKey_(value);
      if (alias && aliases.indexOf(alias) === -1) {
        aliases.push(alias);
      }
    });
  }

  const workDays = new Set();
  if (Array.isArray(opts.workDays)) {
    opts.workDays.forEach(num => {
      const day = Number(num);
      if (!isNaN(day) && day >= 0 && day <= 6) {
        workDays.add(day);
      }
    });
  }
  if (!workDays.size) {
    for (let i = 0; i < 7; i++) workDays.add(i);
  }

  const displayName = opts.displayName || (normalizedId ? normalizedId.split('@')[0] : String(identifier || '')); 

  return {
    id: normalizedId || displayName,
    aliases,
    displayName,
    workDays,
    skipHolidays: !!opts.skipHolidays,
    matches(email){
      const normalized = normalizeEmailKey_(email);
      if (!normalized) return false;
      if (typeof opts.matcher === 'function') {
        try {
          return !!opts.matcher(normalized);
        } catch (err) {
          Logger.log(`[createStaffShiftRule_] matcher failed for ${displayName}: ${err && err.message ? err.message : err}`);
          return false;
        }
      }
      for (let i = 0; i < aliases.length; i++) {
        const alias = aliases[i];
        if (alias && normalized.indexOf(alias) >= 0) {
          return true;
        }
      }
      return false;
    }
  };
}

const STAFF_SHIFT_RULES = [
  createStaffShiftRule_('sugawara@', { displayName: 'sugawara@', workDays: [1,2,3,4,5], skipHolidays: true }),
  createStaffShiftRule_('yanai@', { displayName: 'yanai@', workDays: [1,2,3,4,5], skipHolidays: true }),
  createStaffShiftRule_('nakazawa@', { displayName: 'nakazawa@', workDays: [1,2,3,4,5], skipHolidays: true }),
  createStaffShiftRule_('horiguchi@', { displayName: 'horiguchi@', workDays: [1,2,3,4,5], skipHolidays: true }),
  createStaffShiftRule_('takahiro@', { displayName: 'takahiro@', workDays: [0,1,3,4,6], skipHolidays: true }),
  createStaffShiftRule_('ishimatu@', { displayName: 'ishimatu@', workDays: [0,1,2,3,4], skipHolidays: true }),
  createStaffShiftRule_('maruyama@', { displayName: 'maruyama@', workDays: [1,3,4,5,6], skipHolidays: true }),
  createStaffShiftRule_('takeuti@', { displayName: 'takeuti@', workDays: [1,2,4,6], skipHolidays: true }),
  createStaffShiftRule_('kouno@', { displayName: 'kouno@', workDays: [1,3,5], skipHolidays: true }),
  createStaffShiftRule_('makishima@', { displayName: 'makishima@', workDays: [4,6], skipHolidays: true }),
  createStaffShiftRule_('urano@', { displayName: 'urano@', workDays: [1,2,4,5,6], skipHolidays: true })
];

function resolveStaffDisplayName_(email){
  const normalized = normalizeEmailKey_(email);
  if (!normalized) return '';
  for (let i = 0; i < STAFF_SHIFT_RULES.length; i++) {
    const rule = STAFF_SHIFT_RULES[i];
    if (rule && typeof rule.matches === 'function' && rule.matches(normalized)) {
      return rule.displayName || normalized;
    }
  }
  return normalized.split('@')[0] || normalized;
}

function isJapaneseHoliday_(date){
  if (!(date instanceof Date) || isNaN(date.getTime())) return false;
  try {
    const cal = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    if (!cal) return false;
    const events = cal.getEventsForDay(date);
    return Array.isArray(events) && events.length > 0;
  } catch (err) {
    Logger.log(`[isJapaneseHoliday_] failed: ${err && err.message ? err.message : err}`);
    return false;
  }
}

function isStaffScheduledForDay_(rule, weekday, isHoliday){
  if (!rule) return false;
  if (rule.skipHolidays && isHoliday) {
    return false;
  }
  if (rule.workDays && rule.workDays.size) {
    return rule.workDays.has(weekday);
  }
  return true;
}

function collectTreatmentStaffEmails_(start, end){
  const result = new Set();
  const sheet = sh('施術録');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;
  const width = Math.min(TREATMENT_SHEET_HEADER.length, sheet.getMaxColumns());
  const values = sheet.getRange(2,1,lastRow-1,width).getValues();
  values.forEach(row => {
    const ts = row[0];
    const email = normalizeEmailKey_(row[3]);
    const category = width >= 8 ? String(row[7] || '').trim() : '';
    if (!email || !category) return;
    const when = ts instanceof Date ? ts : new Date(ts);
    if (isNaN(when.getTime()) || when < start || when >= end) return;
    result.add(email);
  });
  return result;
}

function hasRecordedForRule_(rule, recordedEmails){
  if (!recordedEmails || !recordedEmails.size) return false;
  for (const email of recordedEmails) {
    if (rule.matches(email)) {
      return true;
    }
  }
  return false;
}

function checkMissingTreatmentRecords(targetDate){
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const base = targetDate ? new Date(targetDate) : new Date();
  if (isNaN(base.getTime())) {
    throw new Error('日付指定が不正です');
  }

  const start = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  const end = new Date(start.getTime() + 24 * 60 * 60 * 1000);
  const weekday = start.getDay();
  const holiday = isJapaneseHoliday_(start);

  const recorded = collectTreatmentStaffEmails_(start, end);
  const scheduled = STAFF_SHIFT_RULES.filter(rule => isStaffScheduledForDay_(rule, weekday, holiday));
  const missing = scheduled.filter(rule => !hasRecordedForRule_(rule, recorded));

  const summary = {
    date: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
    weekday,
    isHoliday: holiday,
    scheduledCount: scheduled.length,
    missingCount: missing.length,
    recordedCount: recorded.size,
    scheduledStaff: scheduled.map(rule => rule.displayName),
    missingStaff: missing.map(rule => rule.displayName)
  };

  if (!scheduled.length) {
    Logger.log(`[checkMissingTreatmentRecords] 当日の出勤対象者が見つかりません date=${summary.date} holiday=${holiday}`);
    summary.notified = false;
    return summary;
  }

  if (!missing.length) {
    Logger.log(`[checkMissingTreatmentRecords] 施術記録漏れはありません date=${summary.date}`);
    summary.notified = false;
    return summary;
  }

  const staffLines = missing.map(rule => `・${rule.displayName}`).join('\n');
  const message = `⚠️ 本日の施術録記載がされていません。ご確認ください。\n対象スタッフ:\n${staffLines}`;
  notifyChat_(message);
  summary.notified = true;
  return summary;
}

function runMissingTreatmentAlertJob(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('[runMissingTreatmentAlertJob] ロック取得に失敗しました');
    return null;
  }
  try {
    const result = checkMissingTreatmentRecords();
    Logger.log(`[runMissingTreatmentAlertJob] result=${JSON.stringify(result)}`);
    return result;
  } finally {
    lock.releaseLock();
  }
}

function ensureMissingTreatmentAlertTrigger(){
  const handler = 'runMissingTreatmentAlertJob';
  const triggers = ScriptApp.getProjectTriggers();
  let hasClockTrigger = false;
  triggers.forEach(tr => {
    if (tr.getHandlerFunction() === handler) {
      if (tr.getEventType() === ScriptApp.EventType.CLOCK) {
        hasClockTrigger = true;
      } else {
        ScriptApp.deleteTrigger(tr);
      }
    }
  });
  if (!hasClockTrigger) {
    ScriptApp.newTrigger(handler)
      .timeBased()
      .everyDays(1)
      .atHour(19)
      .create();
    Logger.log('[ensureMissingTreatmentAlertTrigger] 新規トリガーを作成しました (19:00 JST)');
  }
  return true;
}

function fetchPatientNamesMap_(idSet){
  const result = new Map();
  if (!idSet || !idSet.size) return result;
  const infoSheet = sh('患者情報');
  const lastRow = infoSheet.getLastRow();
  if (lastRow < 2) return result;
  const lastCol = infoSheet.getLastColumn();
  const headers = infoSheet.getRange(1,1,1,lastCol).getDisplayValues()[0];
  const colRec = getColFlexible_(headers, LABELS.recNo, PATIENT_COLS_FIXED.recNo, '施術録番号');
  const colName = getColFlexible_(headers, LABELS.name, PATIENT_COLS_FIXED.name, '名前');
  const rows = infoSheet.getRange(2,1,lastRow-1,lastCol).getDisplayValues();
  const needed = new Set(Array.from(idSet).map(normId_).filter(Boolean));
  rows.forEach(row => {
    const pid = normId_(row[colRec-1]);
    if (!pid || !needed.has(pid)) return;
    if (!result.has(pid)) {
      result.set(pid, row[colName-1] || '');
    }
  });
  return result;
}

function sendDailySummaryToChat(targetDate){
  ensureAuxSheets_();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const base = targetDate ? new Date(targetDate) : new Date();
  if (isNaN(base.getTime())) {
    throw new Error('日付指定が不正です');
  }
  const start = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  const end = new Date(start.getTime() + 24 * 60 * 60 * 1000);
  const targetDayKey = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
  const sheet = sh('施術録');
  const lastRow = sheet.getLastRow();
  const width = Math.min(TREATMENT_SHEET_HEADER.length, sheet.getMaxColumns());
  const summary = {
    date: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
    staffProcessed: 0,
    posted: 0,
    skipped: 0,
    totalTreatments: 0,
    errors: []
  };
  if (lastRow < 2) {
    Logger.log('[sendDailySummaryToChat] 施術録にデータがありません');
    return summary;
  }

  const values = sheet.getRange(2,1,lastRow-1,width).getValues();
  const byStaff = new Map();
  const patientIds = new Set();

  values.forEach(row => {
    const ts = row[0];
    const rawId = row[1];
    const emailRaw = String(row[3] || '').trim();
    const category = width >= 8 ? String(row[7] || '').trim() : '';
    if (!emailRaw || !category) return;

    const when = ts instanceof Date ? ts : new Date(ts);
    if (isNaN(when.getTime())) return;
    const dayKey = Utilities.formatDate(when, tz, 'yyyy-MM-dd');
    if (dayKey !== targetDayKey) return;

    const key = emailRaw.toLowerCase();
    const entry = byStaff.get(key) || { email: emailRaw, count: 0, patientIds: new Set(), recordNames: new Set(), extras: new Set() };
    entry.count += 1;
    summary.totalTreatments += 1;

    const normalizedId = normId_(rawId);
    if (normalizedId) {
      entry.patientIds.add(normalizedId);
      patientIds.add(normalizedId);
    } else if (rawId) {
      entry.extras.add('ID:' + String(rawId).trim());
    }

    const recordedName = String(row[5] || '').trim();
    if (recordedName) {
      entry.recordNames.add(recordedName);
    }

    byStaff.set(key, entry);
  });

  if (!byStaff.size) {
    Logger.log('[sendDailySummaryToChat] 当日に該当する施術がありません');
    return summary;
  }

  summary.staffProcessed = byStaff.size;

  const nameMap = fetchPatientNamesMap_(patientIds);
  const { map: webhookMap, defaultUrl } = getWebhookConfig_();
  const dateDisp = Utilities.formatDate(start, tz, 'M月d日');

  byStaff.forEach((entry, key) => {
    const webhookUrl = webhookMap.get(key) || defaultUrl;
    const names = new Set();

    entry.patientIds.forEach(pid => {
      const name = nameMap.get(pid);
      if (name) {
        names.add(name);
      }
    });

    entry.recordNames.forEach(name => {
      if (name) names.add(name);
    });

    if (!names.size) {
      entry.patientIds.forEach(pid => names.add('ID:' + pid));
    }

    entry.extras.forEach(label => {
      if (label) names.add(label);
    });

    if (!names.size) {
      names.add('該当なし');
    }

    const nameList = Array.from(names)
      .map(label => {
        const text = String(label || '').trim();
        if (!text) return '';
        if (text.startsWith('ID:') || text.endsWith('様') || text === '該当なし') return text;
        return `${text} 様`;
      })
      .filter(Boolean);
    const message = `本日の施術確認\n${dateDisp} に ${entry.count}件の施術を記録しました。\n患者:\n${nameList.join('\n')}`;

    if (!webhookUrl) {
      Logger.log(`[sendDailySummaryToChat] Webhook未設定 staff=${entry.email}`);
      summary.skipped += 1;
      return;
    }

    try {
      const response = postJsonWebhook_(webhookUrl, { text: message });
      if (response && response.status) {
        Logger.log(`[sendDailySummaryToChat] webhook response staff=${entry.email} status=${response.status}`);
      }
      summary.posted += 1;
    } catch (err) {
      const errMsg = `[sendDailySummaryToChat] 送信失敗 staff=${entry.email} err=${err}`;
      Logger.log(errMsg);
      summary.errors.push(errMsg);
      summary.skipped += 1;
    }
  });

  if (summary.errors.length) {
    const adminWebhook = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL_ADMIN') || defaultUrl;
    if (adminWebhook) {
      try {
        postJsonWebhook_(adminWebhook, { text: `施術記録通知の送信に失敗しました\n${summary.errors.join('\n')}` });
      } catch (err) {
        Logger.log(`[sendDailySummaryToChat] 管理者への通知に失敗 err=${err}`);
      }
    }
  }

  return summary;
}

function runDailySummaryJob(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('[runDailySummaryJob] ロック取得に失敗しました');
    return null;
  }
  try {
    const summary = sendDailySummaryToChat();
    Logger.log(`[runDailySummaryJob] summary=${JSON.stringify(summary)}`);
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function ensureDailySummaryTrigger(){
  const handler = 'runDailySummaryJob';
  const triggers = ScriptApp.getProjectTriggers();
  let hasClockTrigger = false;
  triggers.forEach(tr => {
    if (tr.getHandlerFunction() === handler) {
      if (tr.getEventType() === ScriptApp.EventType.CLOCK) {
        hasClockTrigger = true;
      } else {
        ScriptApp.deleteTrigger(tr);
      }
    }
  });
  if (!hasClockTrigger) {
    ScriptApp.newTrigger(handler)
      .timeBased()
      .everyDays(1)
      .atHour(19)
      .create();
    Logger.log('[ensureDailySummaryTrigger] 新規トリガーを作成しました (19:00 JST)');
  }
  return true;
}
/*** ── Index（ダッシュボード）再構築 ───────────────── **/
function DashboardIndex_refreshAll(){
  ensureAuxSheets_();
  const idx = sh('ダッシュボード'); idx.clearContents();
  idx.getRange(1,1,1,11).setValues([[
    '患者ID','氏名','同意年月日','次回期限','期限ステータス',
    '担当者(60d)','最終施術日','年次要確認','休止','ミュート解除予定','負担割合整合'
  ]]);

  // 患者情報を全件読み
  const sp = sh('患者情報');
  const plc = sp.getLastColumn(), plr = sp.getLastRow();
  if (plr < 2) return;
  const pHead = sp.getRange(1,1,1,plc).getDisplayValues()[0];
  const pvals = sp.getRange(2,1,plr-1,plc).getValues();

  // 施術録から直近60日 担当メール頻度 & 最終施術日
  const rec = sh('施術録');
  const rlr = rec.getLastRow();
  const staffFreqById = new Map();
  const lastVisitById = new Map();
  if (rlr >= 2){
    const rvals = rec.getRange(2,1,rlr-1,6).getValues(); // [TS,施術録番号,所見,メール,最終確認,名前]
    const since = new Date(); since.setDate(since.getDate()-60);
    rvals.forEach(r=>{
      const ts = r[0], id = String(r[1]||'').trim(); if (!id) return;
      const d = ts instanceof Date ? ts : new Date(ts);
      if (isNaN(d.getTime())) return;
      // 最終施術
      const cur = lastVisitById.get(id);
      if (!cur || d > cur) lastVisitById.set(id, d);
      // 直近60日スタッフ頻度
      if (d >= since){
        const mail = String(r[3]||'').trim();
        const m = staffFreqById.get(id) || new Map();
        m.set(mail, (m.get(mail)||0)+1);
        staffFreqById.set(id, m);
      }
    });
  }
  const topFreq = (m)=>{ let best='',n=-1; m&&m.forEach((v,k)=>{ if(v>n){n=v;best=k;} }); return best; };

  // News用の年次要確認（7–8月のみtrue）
  const isAnnualSeason = (()=>{ const mm=(new Date()).getMonth()+1; return (mm===7||mm===8); })();

  // ヘッダ列解決
  const cRec  = getColFlexible_(pHead, LABELS.recNo,  PATIENT_COLS_FIXED.recNo,  '施術録番号');
  const cName = getColFlexible_(pHead, LABELS.name,   PATIENT_COLS_FIXED.name,   '名前');
  const cCons = getColFlexible_(pHead, LABELS.consent,PATIENT_COLS_FIXED.consent,'同意年月日');
  const cShare= getColFlexible_(pHead, LABELS.share,  PATIENT_COLS_FIXED.share,  '負担割合');

  // フラグ（休止/中止/ミュート解除予定）
  const statusOf = (pid)=> getStatus_(pid); // 既存関数を活用

  // 出力行を構築
  const out = pvals.map(r=>{
    const pid   = normId_(r[cRec-1]);
    if (!pid) return null;
    const name  = r[cName-1] || '';
    const cons  = r[cCons-1] || '';
    const next  = calcConsentExpiry_(cons) || '';
    // 期限ステ
    let due = 'ok';
    if (next){
      const n = new Date(next), today = new Date();
      const diff = Math.floor((n - today)/86400000);
      if (diff < 0) due = 'overdue';
      else if (diff <= 14) due = 'nearing';
    }
    const stat = statusOf(pid);
    const staff60 = topFreq(staffFreqById.get(pid));
    const lastV = lastVisitById.get(pid) ? Utilities.formatDate(lastVisitById.get(pid), Session.getScriptTimeZone()||'Asia/Tokyo', 'yyyy-MM-dd') : '';
    const shareRaw = r[cShare-1];
    const shareOk = (shareRaw===1 || shareRaw===2 || shareRaw===3);

    return [pid, name, cons, next, due, staff60, lastV, !!isAnnualSeason, stat.status==='suspended', stat.pauseUntil||'', !!shareOk];
  }).filter(Boolean);

  if (out.length) idx.getRange(2,1,out.length,out[0].length).setValues(out);
}

/** 後で差分化するフック（まずは全件でOK） */
function DashboardIndex_updatePatients(_patientIds){ DashboardIndex_refreshAll(); }
/*** ── 読み取りAPI：getAdminDashboard ───────────── **/
function getAdminDashboard(payload){
  // 1) 権限（社内ドメイン＆管理者判定：ALLOWED_DOMAINが未設定ならスキップ）
  assertDomain_();
  // 代表admin判定は「通知設定.管理者=TRUE」を見る
  if (!isAdminUser_()) throw new Error('管理者権限が必要です');

  // 2) キャッシュ
  const cache = CacheService.getScriptCache();
  const key = 'admin:'+ Utilities.base64EncodeWebSafe(JSON.stringify(payload||{})).slice(0,64);
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  // 3) Index（ダッシュボード）から読み出し
  const idx = sh('ダッシュボード');
  const lr = idx.getLastRow(); if (lr < 2) { DashboardIndex_refreshAll(); } // 初回空なら構築
  const lr2 = idx.getLastRow(); if (lr2 < 2) return { kpi:{}, nearing:[], annual:[], paused:[], invalid:[], serverTime:new Date().toISOString() };

  const vals = idx.getRange(2,1,lr2-1,11).getDisplayValues();
  const head = idx.getRange(1,1,1,11).getDisplayValues()[0];
  const col = Object.fromEntries(head.map((h,i)=>[h,i]));

  // 4) フィルタ適用（段階導入：最初は期間/担当者/ステを無視してOK）
  const nearing = vals.filter(r => String(r[col['期限ステータス']])==='nearing');
  const overdue = vals.filter(r => String(r[col['期限ステータス']])==='overdue');
  const annual  = vals.filter(r => String(r[col['年次要確認']])==='TRUE');
  const paused  = vals.filter(r => String(r[col['休止']])==='TRUE');
  const invalid = vals.filter(r => String(r[col['負担割合整合']])!=='TRUE');

  const res = {
    kpi: {
      nearing: nearing.length,
      overdue: overdue.length,
      annual:  annual.length,
      paused:  paused.length
    },
    nearing: nearing.concat(overdue), // 一覧は“期限接近/超過”をまとめて返す
    annual, paused, invalid,
    serverTime: new Date().toISOString()
  };

  cache.put(key, JSON.stringify(res), 90); // TTL 90s
  return res;
}

function isAdminUser_(){
  const me = normalizeEmailKey_((Session.getActiveUser()||{}).getEmail());
  if (!me) return false;
  if (PAYROLL_OWNER_EMAILS.some(email => normalizeEmailKey_(email) === me)) {
    return true;
  }
  try{
    const s = sh('通知設定');
    const lr = s.getLastRow();
    if (lr < 2) return false;
    const vals = s.getRange(2,1,lr-1,3).getDisplayValues(); // [スタッフメール,WebhookURL,管理者]
    return vals.some(r => {
      const rowEmail = normalizeEmailKey_(r && r[0]);
      const isAdmin = String(r && r[2] || '').trim().toUpperCase() === 'TRUE';
      return rowEmail && rowEmail === me && isAdmin;
    });
  }catch(e){
    return false;
  }
}
/*** ── 書き込みAPI：runBulkActions ───────────── **/
function runBulkActions(actions){
  assertDomain_();
  if (!isAdminUser_()) throw new Error('管理者権限が必要です');
  if (!Array.isArray(actions)||!actions.length) return { ok:true, updated:0 };

  const lock = LockService.getScriptLock(); lock.tryLock(5000);
  try{
    const touched = new Set();
    actions.forEach(a=>{
      const pid = a.patientId; if(!pid) return;
      switch(a.type){
        case 'confirm':      // 同意日 = 今日
          updateConsentDate(pid, Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'Asia/Tokyo','yyyy-MM-dd'));
          touched.add(pid);
          break;
        case 'normalize':    // 負担割合 1/2/3
          updateBurdenShare(pid, String(a.value)); touched.add(pid);
          break;
        case 'unpause':      // 休止解除（= active に）
          // 既存は markSuspend/markStop なので解除ユーティリティを簡便実装
          unpause_(pid); touched.add(pid);
          break;
        case 'annual_ok':    // 年次確認登録
          sh('年次確認').appendRow([String(pid), (a.year||new Date().getFullYear()), new Date(), (Session.getActiveUser()||{}).getEmail() ]);
          pushNews_(pid,'年次確認','年次確認を登録');
          touched.add(pid);
          break;
        case 'schedule':     // 予定登録
          if (a.date){
            sh('予定').appendRow([String(pid),'通院', a.date, (Session.getActiveUser()||{}).getEmail()]);
            pushNews_(pid,'予定','通院予定を登録：'+a.date);
            touched.add(pid);
          }
          break;
      }
    });

    // Index差分更新（v1は全件でOK）
    if (touched.size) {
      const ids = Array.from(touched);
      DashboardIndex_updatePatients(ids);
      invalidatePatientCaches_(ids);
    }
    return { ok:true, updated: actions.length };
  } finally {
    lock.releaseLock();
  }
}

// 休止解除（簡易）
function unpause_(pid){
  const s=sh('フラグ'); s.appendRow([String(pid),'active','']);
  pushNews_(pid,'状態','休止解除');
  log_('休止解除', pid, '');
  invalidatePatientCaches_(pid, { header: true });
}
/*** ── 施術録：タイムスタンプ編集 ───────────────── **/
function updateTreatmentTimestamp(row, newLocal){
  assertDomain_(); ensureAuxSheets_();
  const s = sh('施術録');
  const lr = s.getLastRow();
  if (row <= 1 || row > lr) throw new Error('行が不正です');
  if (!newLocal) throw new Error('日時が空です');

  // 現在の値を退避（監査ログ用）
  const oldTs = s.getRange(row, 1).getValue();        // 列A: タイムスタンプ
  const pid   = String(s.getRange(row, 2).getValue()); // 列B: 施術録番号（患者ID）
  const treatmentId = String(s.getRange(row, 7).getValue() || '').trim();

  // 入力（例: "2025-09-04T14:30" / "2025-09-04 14:30" / "2025/9/4 14:30"）を Date に変換
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const d = parseDateTimeFlexible_(newLocal, tz);
  if (!d || isNaN(d.getTime())) throw new Error('日時の形式が不正です');

  // 書き換え
  s.getRange(row, 1).setValue(d);

  // 監査ログ
  const toDisp = (v)=> v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm') : String(v||'');
  log_('施術TS修正', pid, `row=${row}  ${toDisp(oldTs)} -> ${toDisp(d)}`);
  const newsMeta = treatmentId ? { source: 'treatment', treatmentId } : null;
  pushNews_(pid, '記録', `施術記録の日時を修正: ${toDisp(d)}`, newsMeta);

  // ダッシュボードの最終施術日に影響するので Index を更新（v1は全件でOK）
  DashboardIndex_updatePatients([pid]);

  invalidatePatientCaches_(pid, { header: true, treatments: true, latestTreatmentRow: true });
  return true;
}
/** 文字列→Date（datetime-localや各種区切りに耐性） */
function parseDateTimeFlexible_(input, tz){
  if (input instanceof Date && !isNaN(input.getTime())) return input;
  let s = String(input).trim();
  if (!s) return null;

  // "YYYY-MM-DDTHH:mm" → "YYYY-MM-DD HH:mm"
  s = s.replace('T', ' ');

  // 秒が無ければ付与
  const m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (m) {
    const Y = Number(m[1]), Mo = Number(m[2]) - 1, D = Number(m[3]);
    const h = Number(m[4]||'0'), mi = Number(m[5]||'0'), se = Number(m[6]||'0');
    return new Date(Y, Mo, D, h, mi, se);
  }

  // 素直にDateに投げる（最後の手段）
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function normalizeAutoVitalText_(value) {
  if (value == null) return '';
  if (typeof value === 'string') return value.trim();
  if (Array.isArray(value)) {
    return value
      .map(item => normalizeAutoVitalText_(item))
      .filter(Boolean)
      .join('\n')
      .trim();
  }
  if (typeof value === 'object') {
    const preferredKeys = ['note', 'text', 'body', 'message', 'vitals', 'value'];
    for (let i = 0; i < preferredKeys.length; i += 1) {
      const key = preferredKeys[i];
      if (key in value) {
        const resolved = normalizeAutoVitalText_(value[key]);
        if (resolved) return resolved;
      }
    }

    const merged = Object.keys(value || {})
      .map(k => normalizeAutoVitalText_(value[k]))
      .filter(Boolean)
      .join('\n')
      .trim();
    if (merged) return merged;
    return '';
  }
  return String(value).trim();
}

function genVitals(payload) {
  const randomInt = (min, max) => {
    const lower = Math.ceil(min);
    const upper = Math.floor(max);
    return Math.floor(Math.random() * (upper - lower + 1)) + lower;
  };

  const systolic = randomInt(105, 160);
  let diastolic = randomInt(65, 95);
  if (diastolic >= systolic) {
    diastolic = Math.max(65, systolic - randomInt(10, 20));
  }
  const pulse = randomInt(60, 95);
  const spo2 = randomInt(96, 99);
  const temperature = Math.random() * (36.8 - 36.0) + 36.0;

  const formattedTemp = temperature.toFixed(1);
  return `vital ${systolic}/${diastolic}/${pulse}bpm / SpO2:${spo2}% ${formattedTemp}℃`;
}

function tryGenerateAutoVitals_(payload) {
  try {
    if (typeof genVitals !== 'function') return '';
    const raw = genVitals(payload);
    return normalizeAutoVitalText_(raw);
  } catch (err) {
    const message = err && err.stack ? err.stack : (err && err.message) ? err.message : String(err);
    Logger.log(`[submitTreatment] genVitals() failed: ${message}`);
    return '';
  }
}

function logSubmitTreatmentTimings_(pid, treatmentId, status, timings){
  if (!timings || !timings.length) return;
  const parts = timings.join(' | ');
  Logger.log(`[submitTreatment][${status}] pid=${pid || ''} tid=${treatmentId || ''} ${parts}`);
}

function submitTreatment(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error('保存処理が混み合っています。数秒後に再度お試しください。');
  }
  const startMs = Date.now();
  const timings = [];
  const markTiming = label => { timings.push(`${label}:${Date.now() - startMs}ms`); };
  let pid = '';
  let treatmentIdForLog = String(payload?.treatmentId || '').trim();
  let timingLogged = false;
  try {
    ensureAuxSheets_();
    markTiming('prepared');
    const s = sh('施術録');
    const categoryInfo = resolveTreatmentCategoryFromPayload_(payload);
    const categoryKey = categoryInfo.key || '';
    const categoryLabel = categoryInfo.label || '';

    const presetLabel = String(payload?.presetLabel || '').trim();
    const actions = payload?.actions || {};
    const allowsSameDayUpdate = (
      presetLabel === '同意書受渡'
        || presetLabel === '再同意取得確認'
        || actions.visitPlanDate
        || actions.consentUndecided
    );
    const categoryRequired = !(
      presetLabel === '同意書受渡'
        || presetLabel === '再同意取得確認'
        || actions.visitPlanDate
        || actions.consentUndecided
    );

    if (categoryRequired && !categoryKey) {
      throw new Error('施術区分を特定できませんでした。画面を再読み込みしてから再度お試しください。');
    }
    pid = String(payload?.patientId || '').trim();
    if (pid && categoryKey === 'new') {
      throw new Error('「新規」区分では患者IDを空のまま保存してください。');
    }
    if (!pid && !categoryInfo.allowEmptyPatientId) {
      const labelText = categoryLabel || '施術録';
      throw new Error(`${labelText}を保存するには患者IDを入力してください。`);
    }

    const user = (Session.getActiveUser() || {}).getEmail() || '';

    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    const nowDate = new Date();
    const now = Utilities.formatDate(nowDate, tz, 'yyyy-MM-dd HH:mm:ss');
    markTiming('context');

    const note = String(payload?.notesParts?.note || '').trim();
    let merged = note;
    if (!merged) {
      const autoVitals = tryGenerateAutoVitals_(payload);
      merged = autoVitals || 'バイタル自動記録';
    }
    markTiming('noteReady');

    const incomingTreatmentId = String(payload?.treatmentId || '').trim();
    if (incomingTreatmentId) {
      const dupRow = findTreatmentRowById_(s, incomingTreatmentId);
      markTiming('idCheck');
      if (dupRow) {
        logSubmitTreatmentTimings_(pid, incomingTreatmentId, 'duplicate-id', timings);
        timingLogged = true;
        return {
          ok: false,
          skipped: true,
          duplicate: true,
          msg: '同じ操作が既に保存されています',
          row: dupRow.row,
          treatmentId: incomingTreatmentId,
        };
      }
    }
    if (!incomingTreatmentId) {
      markTiming('idCheck');
    }

    const recentDup = detectRecentDuplicateTreatment_(s, pid, merged, nowDate, tz, incomingTreatmentId);
    markTiming('duplicateScan');
    if (recentDup) {
      if (recentDup.reason === 'recentContent') {
        pushNews_(pid, '警告', '二重登録を検出し保存をスキップしました');
      }
      const duplicateId = recentDup.treatmentId || incomingTreatmentId;
      logSubmitTreatmentTimings_(pid, duplicateId, 'duplicate-content', timings);
      timingLogged = true;
      return {
        ok: false,
        skipped: true,
        duplicate: true,
        msg: recentDup.message,
        row: recentDup.row,
        treatmentId: recentDup.treatmentId,
      };
    }

    const existingToday = findExistingTreatmentOnDate_(s, pid, nowDate, tz, incomingTreatmentId);
    markTiming('sameDayScan');
    if (existingToday) {
      if (allowsSameDayUpdate) {
        const resolvedTreatmentId = existingToday.treatmentId || incomingTreatmentId || Utilities.getUuid();
        treatmentIdForLog = resolvedTreatmentId;
        const existingNoteRaw = Array.isArray(existingToday.row) ? existingToday.row[2] : '';
        const normalizedExistingNote = normalizeTreatmentNoteForComparison_(existingNoteRaw);
        const normalizedIncomingNote = normalizeTreatmentNoteForComparison_(merged);
        let updatedNote = String(existingNoteRaw || '').trim();
        if (normalizedIncomingNote && normalizedExistingNote.indexOf(normalizedIncomingNote) === -1) {
          updatedNote = updatedNote ? `${updatedNote}\n${merged}` : merged;
        }
        if (!existingToday.treatmentId && resolvedTreatmentId) {
          try {
            s.getRange(existingToday.rowNumber, 7).setValue(resolvedTreatmentId);
          } catch (err) {
            Logger.log('[submitTreatment] Failed to write treatmentId to existing row: ' + (err && err.message ? err.message : err));
          }
        }
        const updateResult = updateTreatmentRow(existingToday.rowNumber, updatedNote);
        markTiming('sameDayUpdate');

        const job = { treatmentId: resolvedTreatmentId, treatmentTimestamp: existingToday.row ? existingToday.row[0] : now };
        let hasFollowUp = false;
        if (pid) {
          job.patientId = pid;
        }
        if (categoryKey) {
          job.treatmentCategoryKey = categoryKey;
        }
        if (categoryLabel) {
          job.treatmentCategoryLabel = categoryLabel;
        }
        if (presetLabel) {
          job.presetLabel = presetLabel;
          hasFollowUp = true;
        }

        const burdenShare = payload?.burdenShare;
        if (burdenShare != null && String(burdenShare).trim() !== '') {
          job.burdenShare = String(burdenShare).trim();
          hasFollowUp = true;
        }

        const visitPlanDate = payload?.actions?.visitPlanDate;
        if (visitPlanDate) {
          job.visitPlanDate = String(visitPlanDate).trim();
          if (job.visitPlanDate) {
            hasFollowUp = true;
          } else {
            delete job.visitPlanDate;
          }
        }

        if (payload?.actions && payload.actions.consentUndecided) {
          job.consentUndecided = true;
          hasFollowUp = true;
        }

        if (hasFollowUp && pid) {
          queueAfterTreatmentJob(job);
          markTiming('queueJob');
        } else if (hasFollowUp && !pid) {
          Logger.log('[submitTreatment] Follow-up skipped because patientId is empty');
        }

        markTiming('done');
        logSubmitTreatmentTimings_(pid, resolvedTreatmentId, 'updated-existing', timings);
        timingLogged = true;

        return {
          ok: true,
          updatedRow: updateResult && updateResult.updatedRow ? updateResult.updatedRow : existingToday.rowNumber,
          treatmentId: resolvedTreatmentId,
          updatedExisting: true
        };
      }
      logSubmitTreatmentTimings_(
        pid,
        existingToday.treatmentId || incomingTreatmentId,
        'duplicate-day',
        timings
      );
      timingLogged = true;
      return {
        ok: false,
        skipped: true,
        duplicate: true,
        msg: '本日はすでに施術記録が登録されています。編集が必要な場合は、既存の記録を編集してください。',
        row: existingToday.rowNumber,
        treatmentId: existingToday.treatmentId,
      };
    }

    const treatmentId = incomingTreatmentId || Utilities.getUuid();
    treatmentIdForLog = treatmentId;
    const treatmentCategoryLabel = categoryLabel;
    const treatmentCategoryKey = categoryKey;
    const attendanceMetrics = resolveTreatmentAttendanceMetrics_(categoryInfo);
    const row = [
      now,
      pid,
      merged,
      user,
      '',
      '',
      treatmentId,
      treatmentCategoryLabel,
      attendanceMetrics.convertedCount,
      attendanceMetrics.newPatientCount,
      attendanceMetrics.totalCount,
      ''
    ];
    s.appendRow(row);
    markTiming('appendRow');

    const job = { treatmentId, treatmentTimestamp: now };
    if (pid) {
      job.patientId = pid;
    }
    if (treatmentCategoryKey) {
      job.treatmentCategoryKey = treatmentCategoryKey;
    }
    if (treatmentCategoryLabel) {
      job.treatmentCategoryLabel = treatmentCategoryLabel;
    }
    let hasFollowUp = false;
    if (presetLabel) {
      job.presetLabel = presetLabel;
      hasFollowUp = true;
    }

    const burdenShare = payload?.burdenShare;
    if (burdenShare != null && String(burdenShare).trim() !== '') {
      job.burdenShare = String(burdenShare).trim();
      hasFollowUp = true;
    }

    const visitPlanDate = payload?.actions?.visitPlanDate;
    if (visitPlanDate) {
      job.visitPlanDate = String(visitPlanDate).trim();
      if (job.visitPlanDate) {
        hasFollowUp = true;
      } else {
        delete job.visitPlanDate;
      }
    }

    if (payload?.actions && payload.actions.consentUndecided) {
      job.consentUndecided = true;
      hasFollowUp = true;
    }

    if (hasFollowUp && pid) {
      queueAfterTreatmentJob(job);
      markTiming('queueJob');
    } else if (hasFollowUp && !pid) {
      Logger.log('[submitTreatment] Follow-up skipped because patientId is empty');
    }

    markTiming('done');
    logSubmitTreatmentTimings_(pid, treatmentId, 'ok', timings);
    timingLogged = true;

    if (pid) {
      invalidatePatientCaches_(pid, { header: true, treatments: true, latestTreatmentRow: true });
    }
    return { ok: true, wroteTo: s.getName(), row, treatmentId };
  } finally {
    lock.releaseLock();
    if (!timingLogged && timings.length) {
      logSubmitTreatmentTimings_(pid, treatmentIdForLog, 'error', timings);
    }
  }
}

function completeConsentHandoutFromNews(payload) {
  const pid = String(payload && payload.patientId || '').trim();
  if (!pid) throw new Error('patientIdが空です');
  const consentUndecided = !!(payload && payload.consentUndecided);
  const visitPlanDate = String(payload && payload.visitPlanDate || '').trim();
  const providedNote = String(payload && payload.note || '').trim();
  const note = providedNote
    || (consentUndecided
      ? '同意書受渡。通院日を確認してください。'
      : (visitPlanDate ? `同意書受渡。（通院予定：${visitPlanDate}）` : '同意書受渡。'));
  const actions = {};
  if (consentUndecided) {
    actions.consentUndecided = true;
  } else if (visitPlanDate) {
    actions.visitPlanDate = visitPlanDate;
  }

  const treatmentPayload = {
    patientId: pid,
    presetLabel: '同意書受渡',
    notesParts: { note },
    actions
  };
  if (payload && payload.treatmentId) {
    treatmentPayload.treatmentId = String(payload.treatmentId);
  }

  const result = submitTreatment(treatmentPayload);
  const newsType = String(payload && payload.newsType || '同意').trim() || '同意';
  const newsMessage = String(payload && payload.newsMessage || '');
  const metaType = normalizeNewsMetaType_(payload && payload.newsMetaType ? payload.newsMetaType : '') || 'handover';
  const rowNumber = payload && typeof payload.newsRow === 'number' ? Number(payload.newsRow) : null;
  const cleared = markNewsClearedByType(pid, newsType, {
    messageContains: newsMessage,
    metaType,
    rowNumber
  });
  const dismissed = dismissHandoverReminder({
    patientId: pid,
    newsType,
    newsMessage,
    newsMetaType: metaType,
    newsRow: rowNumber
  });

  return {
    ok: true,
    result,
    cleared,
    dismissed,
    note,
    actions
  };
}

function completeConsentVerificationFromNews(payload) {
  const pid = String(payload && payload.patientId || '').trim();
  if (!pid) throw new Error('patientIdが空です');
  const visitPlanDate = String(payload && payload.visitPlanDate || '').trim();
  const providedNote = String(payload && payload.note || '').trim();
  const note = providedNote
    || (visitPlanDate
      ? `再同意取得確認（通院予定：${visitPlanDate}）`
      : '再同意取得確認。');

  const treatmentPayload = {
    patientId: pid,
    presetLabel: '再同意取得確認',
    notesParts: { note }
  };

  if (payload && payload.treatmentId) {
    treatmentPayload.treatmentId = String(payload.treatmentId);
  }

  const result = submitTreatment(treatmentPayload);
  const newsType = String(payload && payload.newsType || '同意').trim() || '同意';
  const newsMessage = String(payload && payload.newsMessage || '');
  const metaType = payload && payload.newsMetaType ? String(payload.newsMetaType) : '';
  const rowNumber = payload && typeof payload.newsRow === 'number' ? Number(payload.newsRow) : null;
  const cleared = markNewsClearedByType(pid, newsType, {
    messageContains: newsMessage,
    metaType: metaType,
    rowNumber
  });

  return {
    ok: true,
    result,
    cleared,
    note
  };
}

function normalizeTreatmentNoteForComparison_(value){
  if (value == null) return '';
  const text = Array.isArray(value) ? value.join('\n') : String(value);
  if (!text) return '';
  return text
    .replace(/\r/g, '')
    .split('\n')
    .map(line => {
      const trimmed = line.trim();
      if (!trimmed) return '';
      if (/^vital\s/i.test(trimmed) || trimmed === 'バイタル自動記録') {
        return '[AUTO_VITAL]';
      }
      return trimmed;
    })
    .filter(Boolean)
    .join('\n');
}

function detectRecentDuplicateTreatment_(sheet, pid, note, nowDate, tz, ignoreTreatmentId) {
  const lr = sheet.getLastRow();
  if (lr < 2) return null;

  const rowsToScan = Math.min(lr - 1, 20);
  const startRow = Math.max(2, lr - rowsToScan + 1);
  const values = sheet.getRange(startRow, 1, rowsToScan, 7).getValues();
  const nowMs = nowDate.getTime();
  const windowMs = 60 * 1000; // 1分以内の重複をブロック
  const normalizedNote = normalizeTreatmentNoteForComparison_(note);

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const existingPid = String(row[1] || '').trim();
    if (existingPid !== pid) continue;
    const existingNote = normalizeTreatmentNoteForComparison_(row[2]);
    if (existingNote !== normalizedNote) continue;
    const existingTreatmentId = String(row[6] || '').trim();
    if (ignoreTreatmentId && existingTreatmentId && existingTreatmentId === ignoreTreatmentId) {
      return {
        row,
        treatmentId: existingTreatmentId,
        message: '同じ操作が既に保存されています',
        reason: 'sameRequest',
      };
    }
    const tsDate = normalizeTreatmentTimestamp_(row[0], tz);
    if (!tsDate) continue;
    const diff = nowMs - tsDate.getTime();
    if (diff <= windowMs) {
      return {
        row,
        treatmentId: existingTreatmentId,
        message: '直近1分以内に同じ内容が登録済みのため保存をスキップしました',
        reason: 'recentContent',
      };
    }
    if (diff > windowMs) {
      break;
    }
  }
  return null;
}

function findExistingTreatmentOnDate_(sheet, pid, targetDate, tz, ignoreTreatmentId) {
  const normalizedPid = String(pid || '').trim();
  if (!normalizedPid) return null;

  const lr = sheet.getLastRow();
  if (lr < 2) return null;

  const maxCols = sheet.getMaxColumns();
  const width = Math.min(TREATMENT_SHEET_HEADER.length, maxCols);
  const rowsToScan = Math.min(lr - 1, 200);
  const startRow = Math.max(2, lr - rowsToScan + 1);
  const values = sheet.getRange(startRow, 1, lr - startRow + 1, width).getValues();

  const targetDateStr = Utilities.formatDate(targetDate, tz, 'yyyy-MM-dd');
  const startOfDay = normalizeTreatmentTimestamp_(`${targetDateStr} 00:00:00`, tz);
  const startOfDayMs = startOfDay ? startOfDay.getTime() : null;

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const existingPid = String(row[1] || '').trim();
    if (existingPid !== normalizedPid) continue;

    const tsDate = normalizeTreatmentTimestamp_(row[0], tz);
    if (!tsDate) continue;

    const dateStr = Utilities.formatDate(tsDate, tz, 'yyyy-MM-dd');
    if (dateStr === targetDateStr) {
      const existingTreatmentId = String(row[6] || '').trim();
      if (ignoreTreatmentId && existingTreatmentId && existingTreatmentId === ignoreTreatmentId) {
        continue;
      }
      return { rowNumber: startRow + i, treatmentId: existingTreatmentId, row };
    }

    if (startOfDayMs != null && tsDate.getTime() < startOfDayMs) {
      break;
    }
  }

  return null;
}

function resolveTreatmentCategoryFromPayload_(payload){
  const raw = payload && payload.treatmentCategory;
  const keyCandidates = [];
  if (raw && typeof raw === 'object') {
    if (raw.key != null) keyCandidates.push(String(raw.key).trim());
    if (raw.kind != null) keyCandidates.push(String(raw.kind).trim());
    if (raw.saveKind != null) keyCandidates.push(String(raw.saveKind).trim());
  }
  if (payload && payload.saveKind != null) {
    keyCandidates.push(String(payload.saveKind).trim());
  }
  const normalizedKey = keyCandidates.find(key => key && TREATMENT_CATEGORY_DEFINITIONS[key]);
  const definition = normalizedKey ? TREATMENT_CATEGORY_DEFINITIONS[normalizedKey] : null;
  let label = '';
  if (definition) {
    label = definition.label;
  } else if (raw != null) {
    if (typeof raw === 'string') {
      label = String(raw).trim();
    } else if (typeof raw === 'object') {
      if (raw.label != null) {
        label = String(raw.label).trim();
      } else if (raw.tag != null) {
        label = String(raw.tag).trim();
      }
    }
  }
  return {
    key: normalizedKey || '',
    label,
    allowEmptyPatientId: definition ? definition.allowEmptyPatientId === true : false
  };
}

function resolveTreatmentAttendanceMetrics_(categoryInfo){
  const key = categoryInfo && categoryInfo.key ? String(categoryInfo.key).trim() : '';
  const label = categoryInfo && categoryInfo.label ? String(categoryInfo.label).trim() : '';

  const metricsFromKey = key ? TREATMENT_CATEGORY_ATTENDANCE_METRICS[key] : null;
  let metrics = metricsFromKey;
  if (!metrics && label) {
    const matchedKey = Object.keys(TREATMENT_CATEGORY_DEFINITIONS).find(candidateKey => {
      const definition = TREATMENT_CATEGORY_DEFINITIONS[candidateKey];
      return definition && definition.label === label;
    });
    metrics = matchedKey ? TREATMENT_CATEGORY_ATTENDANCE_METRICS[matchedKey] : null;
  }

  if (!metrics) {
    return { convertedCount: '', newPatientCount: '', totalCount: '' };
  }

  let converted = metrics.convertedCount;
  let newCount = metrics.newPatientCount;

  if (typeof converted === 'string') {
    const parsed = Number(converted);
    converted = Number.isFinite(parsed) ? parsed : '';
  } else if (typeof converted !== 'number' || !Number.isFinite(converted)) {
    converted = '';
  }

  if (typeof newCount === 'string') {
    const parsed = Number(newCount);
    newCount = Number.isFinite(parsed) ? parsed : '';
  } else if (typeof newCount !== 'number' || !Number.isFinite(newCount)) {
    newCount = '';
  }

  const hasConverted = typeof converted === 'number' && Number.isFinite(converted);
  const hasNewCount = typeof newCount === 'number' && Number.isFinite(newCount);

  let resolvedTotal = '';
  if (hasConverted || hasNewCount) {
    const total = (hasConverted ? converted : 0) + (hasNewCount ? newCount : 0);
    resolvedTotal = Number.isFinite(total) ? total : '';
  }

  return {
    convertedCount: hasConverted ? converted : '',
    newPatientCount: hasNewCount ? newCount : '',
    totalCount: resolvedTotal
  };
}

function mapTreatmentCategoryCellToKey_(value){
  const label = String(value || '').trim();
  if (!label) return '';
  if (TREATMENT_CATEGORY_LABEL_TO_KEY[label]) {
    return TREATMENT_CATEGORY_LABEL_TO_KEY[label];
  }
  const normalized = label.replace(/\s+/g, '');
  const matched = Object.keys(TREATMENT_CATEGORY_DEFINITIONS).find(key => {
    const def = TREATMENT_CATEGORY_DEFINITIONS[key];
    if (!def || !def.label) return false;
    return def.label.replace(/\s+/g, '') === normalized;
  });
  return matched || '';
}

function formatMinutesAsTimeText_(minutes){
  if (!Number.isFinite(minutes) || minutes < 0) minutes = 0;
  const total = Math.round(minutes);
  const hours = Math.floor(total / 60);
  const mins = Math.abs(total % 60);
  return String(hours).padStart(2, '0') + ':' + String(mins).padStart(2, '0');
}

function formatDurationText_(minutes){
  if (!Number.isFinite(minutes) || minutes <= 0) return '0時間';
  const total = Math.round(minutes);
  const hours = Math.floor(total / 60);
  const mins = Math.abs(total % 60);
  if (mins === 0) {
    return hours + '時間';
  }
  if (hours === 0) {
    return mins + '分';
  }
  return hours + '時間' + mins + '分';
}

function formatFileSizeShort_(bytes){
  const value = Number(bytes);
  if (!Number.isFinite(value) || value <= 0) return '';
  const units = ['B','KB','MB','GB'];
  let current = value;
  let unitIndex = 0;
  while (current >= 1024 && unitIndex < units.length - 1) {
    current /= 1024;
    unitIndex++;
  }
  const display = unitIndex === 0 ? Math.round(current) : Math.round(current * 10) / 10;
  return display + units[unitIndex];
}

function parseTimeTextToMinutes_(value){
  if (value == null || value === '') return NaN;
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getHours() * 60 + value.getMinutes();
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    if (value >= 0 && value <= 1) {
      return Math.round(value * 24 * 60);
    }
    return Math.round(value);
  }
  const text = String(value).trim();
  if (!text) return NaN;
  const normalized = text.replace(/[時h]/gi, ':').replace(/分/g, '');
  const m = normalized.match(/^(\d{1,2})(?::?(\d{2}))?$/);
  if (m) {
    const h = Number(m[1]);
    const mi = Number(m[2] || '0');
    if (Number.isFinite(h) && Number.isFinite(mi)) {
      return h * 60 + mi;
    }
  }
  const numeric = Number(text);
  if (Number.isFinite(numeric)) {
    return Math.round(numeric);
  }
  return NaN;
}

function resolveTimeTextFromCell_(value, displayValue, tz){
  const timezone = tz || Session.getScriptTimeZone() || 'Asia/Tokyo';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, timezone, 'HH:mm');
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    if (value >= 0 && value <= 1) {
      return formatMinutesAsTimeText_(Math.round(value * 24 * 60));
    }
    return formatMinutesAsTimeText_(value);
  }
  const display = String(displayValue || '').trim();
  if (display) {
    const m = display.match(/^(\d{1,2}):(\d{2})$/);
    if (m) {
      const h = String(m[1]).padStart(2, '0');
      const mi = String(m[2]).padStart(2, '0');
      return h + ':' + mi;
    }
    return display;
  }
  const text = String(value || '').trim();
  if (!text) return '';
  const m = text.match(/^(\d{1,2}):(\d{2})$/);
  if (m) {
    const h = String(m[1]).padStart(2, '0');
    const mi = String(m[2]).padStart(2, '0');
    return h + ':' + mi;
  }
  return text;
}

function formatDateKeyFromValue_(value, tz){
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  if (!text) return '';
  const parsed = new Date(text);
  if (isNaN(parsed.getTime())) return '';
  return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
}

function createDateFromKey_(key){
  const parts = String(key || '').split('-');
  if (parts.length !== 3) return null;
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
  return new Date(year, month - 1, day);
}

function buildVisitAttendanceBreakdown_(counts){
  if (!counts) return '';
  const parts = [];
  const insurance = counts.insurance || 0;
  const self30 = counts.self30 || 0;
  const self60 = counts.self60 || 0;
  const mixed = counts.mixed || 0;
  const newcomer = counts.new || 0;
  if (insurance) {
    parts.push('保険:' + insurance);
  }
  const selfTotal = self30 + self60;
  if (selfTotal) {
    const details = [];
    if (self30) details.push('30=' + self30);
    if (self60) details.push('60=' + self60);
    const detailText = details.length ? '(' + details.join(',') + ')' : '';
    parts.push('自費:' + selfTotal + detailText);
  }
  if (mixed) {
    parts.push('混合:' + mixed);
  }
  if (newcomer) {
    parts.push('新規:' + newcomer);
  }
  return parts.join(' / ');
}

function readVisitAttendanceExistingMap_(sheet, tz){
  const result = new Map();
  const lr = sheet.getLastRow();
  if (lr < 2) return result;
  const width = Math.min(VISIT_ATTENDANCE_SHEET_HEADER.length, sheet.getMaxColumns());
  const rows = sheet.getRange(2, 1, lr - 1, width).getValues();
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const dateKey = formatDateKeyFromValue_(row[0], tz);
    const email = String(row[1] || '').trim().toLowerCase();
    if (!dateKey || !email) continue;
    const key = dateKey + '||' + email;
    const flag = String(row[7] || '').trim().toLowerCase();
    const entry = {
      rowNumber: i + 2,
      auto: flag === VISIT_ATTENDANCE_AUTO_FLAG_VALUE || flag === '1' || flag === '自動',
      rawFlag: flag,
      row
    };
    if (!result.has(key)) {
      result.set(key, entry);
    } else {
      const existing = result.get(key);
      if (existing.auto && !entry.auto) {
        result.set(key, entry);
      } else if (!existing.auto && entry.auto) {
        // keep manual entry preference
      }
    }
  }
  return result;
}

function capVisitAttendanceEndMinutes_(startMinutes, breakMinutes, endMinutes, options){
  const opts = options || {};
  const isHourlyStaff = !!opts.isHourlyStaff;
  const originalEndMinutes = Number.isFinite(endMinutes) ? endMinutes : NaN;
  if (!Number.isFinite(originalEndMinutes)) {
    return { endMinutes, adjusted: false, workMinutes: null, originalEndMinutes: endMinutes };
  }
  if (isHourlyStaff) {
    return { endMinutes: originalEndMinutes, adjusted: false, workMinutes: null, originalEndMinutes };
  }
  const limit = VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES;
  if (originalEndMinutes <= limit) {
    return { endMinutes: originalEndMinutes, adjusted: false, workMinutes: null, originalEndMinutes };
  }
  const safeStart = Number.isFinite(startMinutes) ? startMinutes : VISIT_ATTENDANCE_WORK_START_MINUTES;
  const safeBreak = Number.isFinite(breakMinutes) ? breakMinutes : 0;
  const cappedEnd = limit;
  const workMinutes = Math.max(0, cappedEnd - safeStart - safeBreak);
  return { endMinutes: cappedEnd, adjusted: true, workMinutes, originalEndMinutes };
}

function resolveVisitAttendanceRoundedSource_(source, adjusted, fallback){
  const normalized = String(source || '').trim();
  if (!adjusted) {
    return normalized || String(fallback || '').trim();
  }
  const lowered = normalized.toLowerCase();
  if (lowered === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE || lowered === 'paidleave') {
    return VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE;
  }
  if (lowered === 'autorounded18' || lowered === 'manualrounded18') {
    return normalized;
  }
  if (lowered === VISIT_ATTENDANCE_AUTO_FLAG_VALUE) {
    return 'autoRounded18';
  }
  if (lowered === 'manual') {
    return 'manualRounded18';
  }
  const fallbackLower = String(fallback || '').trim().toLowerCase();
  if (fallbackLower === VISIT_ATTENDANCE_AUTO_FLAG_VALUE || fallbackLower === 'auto') {
    return 'autoRounded18';
  }
  if (fallbackLower === 'manual') {
    return 'manualRounded18';
  }
  if (!normalized) {
    return 'manualRounded18';
  }
  return normalized;
}

function syncVisitAttendance(options){
  ensureVisitAttendanceSheet_();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const treatmentSheet = sh('施術録');
  const width = Math.min(TREATMENT_SHEET_HEADER.length, treatmentSheet.getMaxColumns());
  const lastRow = treatmentSheet.getLastRow();
  const summary = {
    targetedRows: 0,
    appended: 0,
    updated: 0,
    manualSkipped: 0,
    errors: 0
  };
  if (lastRow < 2) {
    return summary;
  }
  const rows = treatmentSheet.getRange(2, 1, lastRow - 1, width).getValues();
  const pending = new Map();
  const flagUpdates = [];

  const ensureNumber = value => {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const text = String(value || '').trim();
    if (!text) return 0;
    const num = Number(text.replace(/,/g, ''));
    return Number.isFinite(num) ? num : 0;
  };

  rows.forEach((row, idx) => {
    const rowNumber = idx + 2;
    const existingFlag = width >= 12 ? String(row[11] || '').trim() : '';
    if (existingFlag) {
      return;
    }
    const email = String(row[3] || '').trim();
    if (!email) {
      flagUpdates.push({ rowNumber, value: '要修正:メール未設定' });
      summary.errors++;
      return;
    }
    const ts = normalizeTreatmentTimestamp_(row[0], tz);
    if (!ts) {
      flagUpdates.push({ rowNumber, value: '要修正:日付不正' });
      summary.errors++;
      return;
    }
    const dateKey = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
    const key = dateKey + '||' + email.toLowerCase();
    const categoryKey = mapTreatmentCategoryCellToKey_(width >= 8 ? row[7] : '');
    const pendingEntry = pending.get(key) || {
      dateKey,
      email,
      totalConverted: 0,
      recordCount: 0,
      counts: { insurance: 0, self30: 0, self60: 0, mixed: 0, new: 0 },
      rowNumbers: []
    };
    pendingEntry.totalConverted += ensureNumber(width >= 11 ? row[10] : 0);
    pendingEntry.recordCount += 1;
    if (categoryKey) {
      const group = TREATMENT_CATEGORY_ATTENDANCE_GROUP[categoryKey];
      if (group === 'insurance') pendingEntry.counts.insurance += 1;
      if (group === 'self') {
        if (categoryKey === 'self30') pendingEntry.counts.self30 += 1;
        else if (categoryKey === 'self60') pendingEntry.counts.self60 += 1;
        else pendingEntry.counts.self30 += 1;
      }
      if (group === 'mixed') {
        pendingEntry.counts.mixed += 1;
        pendingEntry.counts.self30 += 1;
      }
      if (group === 'new') pendingEntry.counts.new += 1;
    }
    pendingEntry.rowNumbers.push(rowNumber);
    pending.set(key, pendingEntry);
    summary.targetedRows += 1;
  });

  if (!pending.size) {
    if (flagUpdates.length) {
      flagUpdates.forEach(update => {
        treatmentSheet.getRange(update.rowNumber, 12).setValue(update.value);
      });
    }
    return summary;
  }

  const attendanceSheet = ensureVisitAttendanceSheet_();
  const existingMap = readVisitAttendanceExistingMap_(attendanceSheet, tz);
  const updates = [];
  const appends = [];

  pending.forEach((entry, key) => {
    let workMinutes = Math.max(0, Math.round(entry.totalConverted * 60));
    const breakMinutes = entry.recordCount >= 7 ? 60 : 0;
    let endMinutes = VISIT_ATTENDANCE_WORK_START_MINUTES + workMinutes + breakMinutes;
    const capResult = capVisitAttendanceEndMinutes_(
      VISIT_ATTENDANCE_WORK_START_MINUTES,
      breakMinutes,
      endMinutes,
      { isHourlyStaff: false }
    );
    if (capResult.adjusted) {
      endMinutes = capResult.endMinutes;
      if (Number.isFinite(capResult.workMinutes)) {
        workMinutes = Math.min(workMinutes, capResult.workMinutes);
      }
    }
    const breakdown = buildVisitAttendanceBreakdown_(entry.counts);
    const dateCell = createDateFromKey_(entry.dateKey) || entry.dateKey;
    const rowValues = [
      dateCell,
      entry.email,
      formatMinutesAsTimeText_(VISIT_ATTENDANCE_WORK_START_MINUTES),
      formatMinutesAsTimeText_(endMinutes),
      formatMinutesAsTimeText_(workMinutes),
      formatMinutesAsTimeText_(breakMinutes),
      breakdown,
      VISIT_ATTENDANCE_AUTO_FLAG_VALUE,
      '',
      '',
      '',
      resolveVisitAttendanceRoundedSource_('auto', capResult.adjusted, 'auto')
    ];

    const existing = existingMap.get(key);
    if (existing && !existing.auto) {
      entry.rowNumbers.forEach(rowNumber => {
        flagUpdates.push({ rowNumber, value: '手動調整済' });
      });
      summary.manualSkipped += entry.rowNumbers.length;
      return;
    }

    if (existing && existing.auto) {
      updates.push({ rowNumber: existing.rowNumber, values: rowValues });
      entry.rowNumbers.forEach(rowNumber => {
        flagUpdates.push({ rowNumber, value: '済' });
      });
      summary.updated += entry.rowNumbers.length;
      return;
    }

    appends.push(rowValues);
    entry.rowNumbers.forEach(rowNumber => {
      flagUpdates.push({ rowNumber, value: '済' });
    });
    summary.appended += entry.rowNumbers.length;
  });

  updates.sort((a, b) => a.rowNumber - b.rowNumber).forEach(update => {
    attendanceSheet.getRange(update.rowNumber, 1, 1, VISIT_ATTENDANCE_SHEET_HEADER.length).setValues([update.values]);
  });

  if (appends.length) {
    const startRow = attendanceSheet.getLastRow() + 1;
    attendanceSheet.getRange(startRow, 1, appends.length, VISIT_ATTENDANCE_SHEET_HEADER.length).setValues(appends);
  }

  flagUpdates.forEach(update => {
    treatmentSheet.getRange(update.rowNumber, 12).setValue(update.value);
  });

  return summary;
}

function toBoolean_(value){
  if (value === true) return true;
  if (value === false) return false;
  const text = String(value || '').trim().toLowerCase();
  return text === 'true' || text === '1' || text === 'yes';
}

function readVisitAttendanceStaffSettings_(){
  const sheet = ensureVisitAttendanceStaffSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Map();
  const width = Math.min(VISIT_ATTENDANCE_STAFF_SHEET_HEADER.length, sheet.getMaxColumns());
  const range = sheet.getRange(2, 1, lastRow - 1, width);
  const values = range.getValues();
  const map = new Map();
  values.forEach(row => {
    const email = normalizeEmailKey_(row[0]);
    if (!email) return;
    const quotaRaw = row[2];
    let quota = Number(quotaRaw);
    if (!Number.isFinite(quota) || quota < 0) {
      const text = String(quotaRaw || '').trim();
      if (text) {
        const parsed = Number(text.replace(/[^0-9.-]/g, ''));
        if (Number.isFinite(parsed)) quota = parsed;
      }
    }
    if (!Number.isFinite(quota) || quota < 0) {
      quota = DEFAULT_ANNUAL_PAID_LEAVE_DAYS;
    }
    const employmentType = normalizeEmploymentType_(row[3]);
    const employmentLabel = getEmploymentLabel_(employmentType);
    const albyteStaffId = String(row[4] || '').trim();
    let defaultShiftMinutes = Number(row[5]);
    if (!Number.isFinite(defaultShiftMinutes) || defaultShiftMinutes <= 0) {
      const defaultText = String(row[5] || '').trim();
      if (defaultText) {
        const parsedDefault = Number(defaultText.replace(/[^0-9.-]/g, ''));
        if (Number.isFinite(parsedDefault) && parsedDefault > 0) {
          defaultShiftMinutes = parsedDefault;
        }
      }
    }
    if (!Number.isFinite(defaultShiftMinutes) || defaultShiftMinutes <= 0) {
      defaultShiftMinutes = null;
    }
    const displayName = String(row[1] || '').trim() || resolveStaffDisplayName_(email) || '';
    const payrollEmployeeId = String(row[6] || '').trim();
    const worksite = String(row[7] || '').trim();
    const worksiteKey = normalizePayrollBaseKey_(worksite);
    map.set(email, {
      email,
      displayName,
      quotaDays: quota,
      employmentType,
      employmentLabel,
      albyteStaffId,
      defaultShiftMinutes,
      payrollEmployeeId,
      worksite,
      worksiteKey
    });
  });
  return map;
}

function normalizeEmploymentType_(value){
  const text = String(value || '').trim().toLowerCase();
  if (!text) return 'employee';
  if (text === 'parttime' || text === 'part-time' || text === 'hourly' || text === 'アルバイト' || text === 'バイト') {
    return 'partTime';
  }
  if (text === 'daily' || text === '日給') {
    return 'daily';
  }
  return 'employee';
}

function getEmploymentLabel_(type){
  const key = String(type || '').toLowerCase();
  if (PAID_LEAVE_EMPLOYMENT_LABELS[key]) {
    return PAID_LEAVE_EMPLOYMENT_LABELS[key];
  }
  return PAID_LEAVE_EMPLOYMENT_LABELS.employee;
}

function normalizePaidLeaveRequestType_(value){
  const text = String(value || '').trim().toLowerCase();
  if (!text) return 'full';
  if (['am','amhalf','am_half','half_am','amhalfday','amhalfday'].indexOf(text) !== -1) {
    return 'amHalf';
  }
  if (['pm','pmhalf','pm_half','half_pm','pmhalfday'].indexOf(text) !== -1) {
    return 'pmHalf';
  }
  if (text === 'half' || text === 'halfday') {
    return 'pmHalf';
  }
  return 'full';
}

function getPaidLeaveTypeLabel_(type){
  const normalized = normalizePaidLeaveRequestType_(type);
  return PAID_LEAVE_TYPE_LABELS[normalized] || PAID_LEAVE_TYPE_LABELS.full;
}

function roundMinutesToIncrement_(value, increment){
  if (!Number.isFinite(value)) return NaN;
  const unit = Number.isFinite(increment) && increment > 0 ? increment : VISIT_ATTENDANCE_ROUNDING_MINUTES;
  return Math.round(value / unit) * unit;
}

function formatMinutesRangeText_(startMinutes, endMinutes){
  if (!Number.isFinite(startMinutes) || !Number.isFinite(endMinutes)) return '';
  return formatMinutesAsTimeText_(startMinutes) + '〜' + formatMinutesAsTimeText_(endMinutes);
}

function resolveVisitAttendanceStaffProfile_(email, options){
  const normalized = normalizeEmailKey_(email);
  const settings = options && options.staffSettings instanceof Map ? options.staffSettings : readVisitAttendanceStaffSettings_();
  if (normalized && settings.has(normalized)) {
    return settings.get(normalized);
  }
  return {
    email: normalized || '',
    displayName: resolveStaffDisplayName_(normalized) || '',
    quotaDays: DEFAULT_ANNUAL_PAID_LEAVE_DAYS,
    employmentType: 'employee',
    employmentLabel: getEmploymentLabel_('employee'),
    albyteStaffId: '',
    defaultShiftMinutes: null,
    payrollEmployeeId: '',
    worksite: '',
    worksiteKey: ''
  };
}

function resolveVisitAttendanceShiftForProfile_(profile, targetDay, options){
  if (!profile || String(profile.employmentType || '').toLowerCase() !== 'parttime') {
    return null;
  }
  const tz = (options && options.tz) || Session.getScriptTimeZone() || 'Asia/Tokyo';
  const context = options && options.shiftContext ? options.shiftContext : readAlbyteShiftRecords_();
  const dateKey = Utilities.formatDate(targetDay, tz, 'yyyy-MM-dd');
  const records = Array.isArray(context && context.records) ? context.records : [];
  let match = null;
  if (profile.albyteStaffId) {
    match = records.find(record => record && record.dateKey === dateKey && record.staffId === profile.albyteStaffId);
  }
  if (!match) {
    const normalizedName = normalizeAlbyteName_(profile.displayName || '');
    if (normalizedName) {
      match = records.find(record => record && record.dateKey === dateKey && record.normalizedName === normalizedName);
    }
  }
  if (!match) {
    return null;
  }
  let startMinutes = Number(match.startMinutes);
  let endMinutes = Number(match.endMinutes);
  if (!Number.isFinite(startMinutes) || !Number.isFinite(endMinutes) || endMinutes <= startMinutes) {
    return null;
  }
  startMinutes = roundMinutesToIncrement_(startMinutes, VISIT_ATTENDANCE_ROUNDING_MINUTES);
  endMinutes = roundMinutesToIncrement_(endMinutes, VISIT_ATTENDANCE_ROUNDING_MINUTES);
  if (!Number.isFinite(startMinutes) || !Number.isFinite(endMinutes) || endMinutes <= startMinutes) {
    return null;
  }
  return {
    startMinutes,
    endMinutes,
    source: 'albyte'
  };
}

function buildPaidLeavePlan_(email, targetDay, options){
  if (!(targetDay instanceof Date) || isNaN(targetDay.getTime())) {
    throw new Error('対象日を解析できませんでした');
  }
  const normalizedEmail = normalizeEmailKey_(email);
  if (!normalizedEmail) {
    throw new Error('スタッフ情報を取得できませんでした');
  }
  const tz = (options && options.tz) || Session.getScriptTimeZone() || 'Asia/Tokyo';
  const staffSettings = options && options.staffSettings instanceof Map ? options.staffSettings : readVisitAttendanceStaffSettings_();
  const profile = resolveVisitAttendanceStaffProfile_(normalizedEmail, { staffSettings });
  const employmentType = String(profile && profile.employmentType || 'employee').toLowerCase();
  let shiftStartMinutes = VISIT_ATTENDANCE_WORK_START_MINUTES;
  let shiftEndMinutes = shiftStartMinutes + PAID_LEAVE_DEFAULT_WORK_MINUTES;
  let shiftSource = 'default';
  if (employmentType === 'parttime') {
    const shift = resolveVisitAttendanceShiftForProfile_(profile, targetDay, { tz, shiftContext: options && options.shiftContext });
    if (shift) {
      shiftStartMinutes = shift.startMinutes;
      shiftEndMinutes = shift.endMinutes;
      shiftSource = shift.source || 'albyte';
    } else if (Number.isFinite(profile.defaultShiftMinutes) && profile.defaultShiftMinutes > 0) {
      shiftEndMinutes = shiftStartMinutes + profile.defaultShiftMinutes;
      shiftSource = 'staffDefault';
    } else {
      throw new Error('この日の予定勤務時間が登録されていません。先にシフトを登録してください。');
    }
  }
  let fullWorkMinutes = employmentType === 'employee'
    ? PAID_LEAVE_DEFAULT_WORK_MINUTES
    : Math.max(0, shiftEndMinutes - shiftStartMinutes);
  if (!Number.isFinite(fullWorkMinutes) || fullWorkMinutes <= 0) {
    fullWorkMinutes = PAID_LEAVE_DEFAULT_WORK_MINUTES;
  }
  const breakMinutes = 0;
  const requestedType = normalizePaidLeaveRequestType_(options && options.requestType);
  const halfWorkMinutes = Math.max(
    VISIT_ATTENDANCE_ROUNDING_MINUTES,
    roundMinutesToIncrement_(fullWorkMinutes / 2, VISIT_ATTENDANCE_ROUNDING_MINUTES)
  );
  const halfAvailable = fullWorkMinutes >= PAID_LEAVE_HALF_MINIMUM_MINUTES;
  let leaveType = requestedType;
  let workMinutes = fullWorkMinutes;
  let startMinutes = shiftStartMinutes;
  let blockedReason = '';
  if ((leaveType === 'amHalf' || leaveType === 'pmHalf')) {
    if (!halfAvailable) {
      blockedReason = 'この日の所定勤務時間では半休を取得できません。';
    }
    workMinutes = Math.min(fullWorkMinutes, halfWorkMinutes);
    if (leaveType === 'amHalf') {
      startMinutes = shiftStartMinutes + Math.max(0, fullWorkMinutes - workMinutes);
    } else {
      startMinutes = shiftStartMinutes;
    }
  }
  workMinutes = roundMinutesToIncrement_(workMinutes, VISIT_ATTENDANCE_ROUNDING_MINUTES);
  startMinutes = roundMinutesToIncrement_(startMinutes, VISIT_ATTENDANCE_ROUNDING_MINUTES);
  if (!Number.isFinite(workMinutes) || workMinutes <= 0) {
    workMinutes = PAID_LEAVE_DEFAULT_WORK_MINUTES;
  }
  if (!Number.isFinite(startMinutes)) {
    startMinutes = VISIT_ATTENDANCE_WORK_START_MINUTES;
  }
  const endMinutes = startMinutes + workMinutes;
  const plan = {
    planVersion: 'paidLeave/v2',
    email: normalizedEmail,
    employmentType,
    employmentLabel: profile && profile.employmentLabel ? profile.employmentLabel : getEmploymentLabel_(employmentType),
    leaveType,
    leaveLabel: getPaidLeaveTypeLabel_(leaveType),
    workMinutes,
    fullWorkMinutes,
    halfWorkMinutes,
    breakMinutes,
    startMinutes,
    endMinutes,
    shiftStartMinutes,
    shiftEndMinutes,
    shiftText: formatMinutesRangeText_(shiftStartMinutes, shiftEndMinutes),
    shiftSource,
    halfAvailable,
    blockedReason
  };
  plan.workText = formatDurationText_(plan.workMinutes);
  plan.fullWorkText = formatDurationText_(plan.fullWorkMinutes);
  plan.halfWorkText = formatDurationText_(plan.halfWorkMinutes);
  return plan;
}

function parsePaidLeaveRequestMetadata_(raw){
  if (!raw && raw !== 0) return null;
  if (typeof raw === 'object' && raw !== null) {
    return raw;
  }
  const text = String(raw || '').trim();
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch (err) {
    return null;
  }
}

function parsePaidLeaveTargetDate_(value){
  let targetDate = value;
  if (targetDate instanceof Date && !isNaN(targetDate.getTime())) {
    return new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
  }
  const raw = String(targetDate || '').trim();
  if (!raw) {
    throw new Error('有給申請の日付を指定してください');
  }
  const normalized = raw.replace(/[\.\/]/g, '-');
  const m = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) {
    throw new Error('日付の形式が不正です (YYYY-MM-DD)');
  }
  const year = Number(m[1]);
  const monthIndex = Number(m[2]) - 1;
  const day = Number(m[3]);
  const parsed = new Date(year, monthIndex, day);
  if (isNaN(parsed.getTime())) {
    throw new Error('日付の解析に失敗しました');
  }
  return parsed;
}

function resolveAnnualPaidLeaveQuota_(email, options){
  const normalized = normalizeEmailKey_(email);
  if (!normalized) return DEFAULT_ANNUAL_PAID_LEAVE_DAYS;
  const settings = options && options.staffSettings instanceof Map ? options.staffSettings : readVisitAttendanceStaffSettings_();
  const entry = settings.get(normalized);
  if (entry && Number.isFinite(entry.quotaDays)) {
    return entry.quotaDays;
  }
  return DEFAULT_ANNUAL_PAID_LEAVE_DAYS;
}

function calculatePaidLeaveUsageForYear_(email, year, tz){
  const normalized = normalizeEmailKey_(email);
  if (!normalized) {
    return { usedDays: 0, records: [] };
  }
  const start = new Date(year, 0, 1);
  const end = new Date(year, 11, 31);
  const records = readVisitAttendanceRecordsForEmail_(normalized, { startDate: start, endDate: end, tz });
  const usedRecords = records.filter(record => (record.leaveType || '').toLowerCase() === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE);
  return { usedDays: usedRecords.length, records: usedRecords };
}

function readVisitAttendanceRecords_(options){
  const opts = options || {};
  const tz = opts.tz || Session.getScriptTimeZone() || 'Asia/Tokyo';
  const startDate = opts.startDate instanceof Date ? new Date(opts.startDate.getFullYear(), opts.startDate.getMonth(), opts.startDate.getDate()) : null;
  const endDate = opts.endDate instanceof Date ? new Date(opts.endDate.getFullYear(), opts.endDate.getMonth(), opts.endDate.getDate()) : null;
  const startMs = startDate ? startDate.getTime() : null;
  const endMs = endDate ? endDate.getTime() : null;
  const emailFilter = normalizeEmailKey_(opts.email);
  const sheet = ensureVisitAttendanceSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const width = Math.min(VISIT_ATTENDANCE_SHEET_HEADER.length, sheet.getMaxColumns());
  const range = sheet.getRange(2, 1, lastRow - 1, width);
  const values = range.getValues();
  const displays = range.getDisplayValues();
  const weekdays = ['日','月','火','水','木','金','土'];
  const staffSettings = opts.staffSettings instanceof Map ? opts.staffSettings : readVisitAttendanceStaffSettings_();
  const results = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const display = displays[i];
    const normalizedEmail = normalizeEmailKey_(row[1] || display[1]);
    if (!normalizedEmail) continue;
    if (emailFilter && normalizedEmail !== emailFilter) continue;

    let dateObj = row[0];
    if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
      const key = formatDateKeyFromValue_(row[0], tz) || formatDateKeyFromValue_(display[0], tz);
      dateObj = createDateFromKey_(key || '');
    }
    if (!dateObj) continue;
    const day = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
    const dayMs = day.getTime();
    if (startMs != null && dayMs < startMs) continue;
    if (endMs != null && dayMs > endMs) continue;

    const startText = resolveTimeTextFromCell_(row[2], display[2], tz);
    const endText = resolveTimeTextFromCell_(row[3], display[3], tz);
    let workText = resolveTimeTextFromCell_(row[4], display[4], tz);
    let breakText = resolveTimeTextFromCell_(row[5], display[5], tz);
    const startMinutes = parseTimeTextToMinutes_(startText);
    const endMinutes = parseTimeTextToMinutes_(endText);
    const originalEndMinutes = Number.isFinite(endMinutes) ? endMinutes : null;
    let workMinutes = parseTimeTextToMinutes_(workText);
    let breakMinutes = parseTimeTextToMinutes_(breakText);
    if (!Number.isFinite(breakMinutes)) breakMinutes = 0;
    if (!Number.isFinite(workMinutes) && Number.isFinite(startMinutes) && Number.isFinite(endMinutes)) {
      workMinutes = Math.max(0, endMinutes - startMinutes - breakMinutes);
      workText = formatMinutesAsTimeText_(workMinutes);
    }
    if (!workText && Number.isFinite(workMinutes)) {
      workText = formatMinutesAsTimeText_(workMinutes);
    }
    if (!breakText && Number.isFinite(breakMinutes)) {
      breakText = formatMinutesAsTimeText_(breakMinutes);
    }

    const breakdown = String(display[6] || row[6] || '').trim();
    const flagRaw = String(display[7] || row[7] || '').trim();
    const auto = flagRaw.toLowerCase() === VISIT_ATTENDANCE_AUTO_FLAG_VALUE || flagRaw === '自動';
    const leaveType = String((row[8] != null && row[8] !== '') ? row[8] : (display[8] != null ? display[8] : '')).trim();
    const hourlyRaw = row[9] != null && row[9] !== '' ? row[9] : display[9];
    const dailyRaw = row[10] != null && row[10] !== '' ? row[10] : display[10];
    const sourceRaw = String((row[11] != null && row[11] !== '') ? row[11] : (display[11] != null ? display[11] : '')).trim();
    const isHourlyStaff = toBoolean_(hourlyRaw);
    const isDailyStaff = toBoolean_(dailyRaw);

    const normalizedSource = sourceRaw.toLowerCase();
    const isRoundedSource = normalizedSource === 'autorounded18' || normalizedSource === 'manualrounded18';
    let effectiveEndMinutes = Number.isFinite(endMinutes) ? endMinutes : null;
    let autoAdjustedEnd = false;
    if (!isHourlyStaff && Number.isFinite(effectiveEndMinutes) && effectiveEndMinutes > VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES) {
      effectiveEndMinutes = VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES;
      autoAdjustedEnd = true;
    }
    if (isRoundedSource) {
      autoAdjustedEnd = true;
      if (!Number.isFinite(effectiveEndMinutes) || effectiveEndMinutes > VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES) {
        effectiveEndMinutes = VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES;
      }
    }
    if (autoAdjustedEnd && Number.isFinite(startMinutes) && Number.isFinite(effectiveEndMinutes)) {
      const cappedWork = Math.max(0, effectiveEndMinutes - startMinutes - breakMinutes);
      if (!Number.isFinite(workMinutes) || workMinutes > cappedWork) {
        workMinutes = cappedWork;
        workText = formatMinutesAsTimeText_(workMinutes);
      }
    }

    let sourceLabel = auto ? '自動反映' : (flagRaw ? flagRaw : '手動入力');
    if (sourceRaw) {
      if (sourceRaw === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE || sourceRaw.toLowerCase() === 'paidleave') {
        sourceLabel = '有給';
      } else if (sourceRaw === 'auto') {
        sourceLabel = '自動反映';
      } else if (sourceRaw === 'manual') {
        sourceLabel = '手動入力';
      } else {
        sourceLabel = sourceRaw;
      }
    }
    if (leaveType === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE) {
      sourceLabel = '有給';
    }
    if (isRoundedSource) {
      sourceLabel = normalizedSource === 'autorounded18' ? '自動反映（18:00調整）' : '手動入力（18:00調整）';
    } else if (autoAdjustedEnd && sourceLabel && sourceLabel.indexOf('18:00調整') === -1) {
      sourceLabel = sourceLabel + '（18:00調整）';
    }

    const finalEndMinutes = Number.isFinite(effectiveEndMinutes)
      ? effectiveEndMinutes
      : (Number.isFinite(endMinutes) ? endMinutes : null);
    const finalEndText = Number.isFinite(finalEndMinutes)
      ? formatMinutesAsTimeText_(finalEndMinutes)
      : (endText || (Number.isFinite(endMinutes) ? formatMinutesAsTimeText_(endMinutes) : ''));
    const autoAdjustmentMessage = autoAdjustedEnd ? '自動補正：退勤は18:00に調整されました' : '';
    const staffSetting = staffSettings.get(normalizedEmail);
    const displayName = staffSetting && staffSetting.displayName
      ? staffSetting.displayName
      : resolveStaffDisplayName_(normalizedEmail);

    results.push({
      email: normalizedEmail,
      staffName: displayName || normalizedEmail,
      date: Utilities.formatDate(day, tz, 'yyyy-MM-dd'),
      displayDate: Utilities.formatDate(day, tz, 'M/d'),
      weekday: weekdays[day.getDay()] || '',
      start: startText || (Number.isFinite(startMinutes) ? formatMinutesAsTimeText_(startMinutes) : ''),
      end: finalEndText,
      work: workText || '',
      break: breakText || '',
      startMinutes: Number.isFinite(startMinutes) ? startMinutes : null,
      endMinutes: Number.isFinite(finalEndMinutes) ? finalEndMinutes : null,
      originalEndMinutes,
      workMinutes: Number.isFinite(workMinutes) ? workMinutes : null,
      breakMinutes: Number.isFinite(breakMinutes) ? breakMinutes : 0,
      breakdown,
      flag: flagRaw,
      auto,
      sourceLabel,
      leaveType,
      isHourlyStaff,
      isDailyStaff,
      source: sourceRaw,
      rowNumber: i + 2,
      autoAdjustedEnd,
      autoAdjustmentMessage
    });
  }
  results.sort((a, b) => {
    const dateDiff = a.date.localeCompare(b.date);
    if (dateDiff !== 0) return dateDiff;
    const nameDiff = (a.staffName || '').localeCompare(b.staffName || '');
    if (nameDiff !== 0) return nameDiff;
    const startDiff = (a.start || '').localeCompare(b.start || '');
    if (startDiff !== 0) return startDiff;
    return (a.email || '').localeCompare(b.email || '');
  });
  return results;
}


function readVisitAttendanceRecordsForEmail_(email, options){
  const normalizedEmail = normalizeEmailKey_(email);
  if (!normalizedEmail) return [];
  const opts = Object.assign({}, options || {}, { email: normalizedEmail });
  return readVisitAttendanceRecords_(opts);
}
function readVisitAttendanceRequests_(options){
  const opts = options || {};
  const tz = opts.tz || Session.getScriptTimeZone() || 'Asia/Tokyo';
  const sheet = ensureVisitAttendanceRequestSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const width = Math.min(VISIT_ATTENDANCE_REQUEST_SHEET_HEADER.length, sheet.getMaxColumns());
  const range = sheet.getRange(2, 1, lastRow - 1, width);
  const values = range.getValues();
  const displays = range.getDisplayValues();
  const normalizedEmail = normalizeEmailKey_(opts.email);
  const statusFilter = opts.status ? (Array.isArray(opts.status) ? opts.status : [opts.status]) : null;
  const statusSet = statusFilter ? new Set(statusFilter.map(v => String(v || '').toLowerCase())) : null;
  const idFilter = opts.id ? String(opts.id).trim() : '';
  const startDate = opts.startDate instanceof Date ? new Date(opts.startDate.getFullYear(), opts.startDate.getMonth(), opts.startDate.getDate()) : null;
  const endDate = opts.endDate instanceof Date ? new Date(opts.endDate.getFullYear(), opts.endDate.getMonth(), opts.endDate.getDate()) : null;
  const startMs = startDate ? startDate.getTime() : null;
  const endMs = endDate ? endDate.getTime() : null;
  const weekdays = ['日','月','火','水','木','金','土'];
  const results = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const display = displays[i];
    const id = String(row[0] || display[0] || '').trim();
    if (idFilter && id !== idFilter) continue;

    const applicantEmail = normalizeEmailKey_(row[2] || display[2]);
    const targetEmail = normalizeEmailKey_(row[3] || display[3] || applicantEmail);
    if (normalizedEmail && targetEmail !== normalizedEmail) continue;

    let targetDate = row[4];
    if (!(targetDate instanceof Date) || isNaN(targetDate.getTime())) {
      const key = formatDateKeyFromValue_(row[4], tz) || formatDateKeyFromValue_(display[4], tz);
      targetDate = createDateFromKey_(key || '');
    }
    if (!targetDate) continue;
    const day = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
    const dayMs = day.getTime();
    if (startMs != null && dayMs < startMs) continue;
    if (endMs != null && dayMs > endMs) continue;

    const statusRaw = String(row[9] || display[9] || 'pending').trim().toLowerCase() || 'pending';
    if (statusSet && !statusSet.has(statusRaw)) continue;

    const createdAt = row[1] instanceof Date && !isNaN(row[1].getTime()) ? row[1] : null;
    const statusUpdatedAt = row[10] instanceof Date && !isNaN(row[10].getTime()) ? row[10] : null;
    const breakMinutes = parseTimeTextToMinutes_(row[7] != null && row[7] !== '' ? row[7] : display[7]);
    const startText = String(row[5] || display[5] || '').trim();
    const endText = String(row[6] || display[6] || '').trim();
    const startMinutes = parseTimeTextToMinutes_(startText);
    const endMinutes = parseTimeTextToMinutes_(endText);

    let originalData = null;
    const originalRaw = row[13] != null && row[13] !== '' ? row[13] : display[13];
    if (originalRaw != null && originalRaw !== '') {
      const text = String(originalRaw);
      try {
        originalData = JSON.parse(text);
      } catch (err) {
        originalData = text;
      }
    }

    const typeRaw = String((row[14] != null && row[14] !== '') ? row[14] : (display[14] != null ? display[14] : '')).trim().toLowerCase();
    const requestType = typeRaw || VISIT_ATTENDANCE_REQUEST_TYPE_CORRECTION;
    let paidLeaveDetail = null;
    if (requestType === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE) {
      const planMeta = parsePaidLeaveRequestMetadata_(row[13] != null && row[13] !== '' ? row[13] : display[13]);
      const leaveType = planMeta && planMeta.leaveType ? planMeta.leaveType : 'full';
      const workMinutes = Number(planMeta && planMeta.workMinutes);
      const shiftText = planMeta && planMeta.shiftText
        ? planMeta.shiftText
        : formatMinutesRangeText_(planMeta && planMeta.shiftStartMinutes, planMeta && planMeta.shiftEndMinutes);
      paidLeaveDetail = {
        leaveType,
        leaveLabel: getPaidLeaveTypeLabel_(leaveType),
        workText: planMeta && planMeta.workText ? planMeta.workText : formatDurationText_(Number.isFinite(workMinutes) && workMinutes > 0 ? workMinutes : PAID_LEAVE_DEFAULT_WORK_MINUTES),
        shiftText: shiftText || '',
        employmentType: planMeta && planMeta.employmentType ? planMeta.employmentType : 'employee',
        employmentLabel: planMeta && planMeta.employmentLabel ? planMeta.employmentLabel : getEmploymentLabel_(planMeta && planMeta.employmentType ? planMeta.employmentType : 'employee')
      };
    }

    results.push({
      id,
      rowNumber: i + 2,
      applicantEmail: applicantEmail || '',
      targetEmail: targetEmail || '',
      targetDate: Utilities.formatDate(day, tz, 'yyyy-MM-dd'),
      targetWeekday: weekdays[day.getDay()] || '',
      monthKey: Utilities.formatDate(day, tz, 'yyyy-MM'),
      createdAt: createdAt ? createdAt.toISOString() : '',
      createdAtText: createdAt ? Utilities.formatDate(createdAt, tz, 'yyyy-MM-dd HH:mm') : String(display[1] || ''),
      start: startText,
      end: endText,
      startMinutes: Number.isFinite(startMinutes) ? startMinutes : null,
      endMinutes: Number.isFinite(endMinutes) ? endMinutes : null,
      breakMinutes: Number.isFinite(breakMinutes) ? breakMinutes : 0,
      breakText: formatMinutesAsTimeText_(Number.isFinite(breakMinutes) ? breakMinutes : 0),
      note: String(row[8] || display[8] || '').trim(),
      status: statusRaw,
      statusLabel: statusRaw === 'approved' ? '承認済み' : statusRaw === 'rejected' ? '差し戻し' : '申請中',
      statusUpdatedAt: statusUpdatedAt ? statusUpdatedAt.toISOString() : '',
      statusUpdatedAtText: statusUpdatedAt ? Utilities.formatDate(statusUpdatedAt, tz, 'yyyy-MM-dd HH:mm') : String(display[10] || ''),
      statusBy: String(row[11] || display[11] || '').trim(),
      statusNote: String(row[12] || display[12] || '').trim(),
      originalData,
      type: requestType,
      paidLeaveDetail,
      typeLabel: requestType === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE ? '有給申請' : '勤怠修正申請'
    });
  }

  results.sort((a, b) => {
    if (a.targetDate === b.targetDate) {
      return (b.createdAt || '').localeCompare(a.createdAt || '');
    }
    return b.targetDate.localeCompare(a.targetDate);
  });
  return results;
}

function buildVisitAttendancePortalMonths_(tz, now, count){
  const list = [];
  const base = new Date(now.getFullYear(), now.getMonth(), 1);
  const total = Math.max(1, Number(count) || 12);
  for (let i = 0; i < total; i++) {
    const date = new Date(base.getFullYear(), base.getMonth() - i, 1);
    list.push({
      key: Utilities.formatDate(date, tz, 'yyyy-MM'),
      label: Utilities.formatDate(date, tz, 'yyyy年M月'),
      requestable: date.getTime() < base.getTime()
    });
  }
  return list;
}

function resolveVisitAttendanceMonthRange_(monthKey, tz, now){
  const currentStart = new Date(now.getFullYear(), now.getMonth(), 1);
  if (monthKey && typeof monthKey === 'string') {
    const text = monthKey.trim();
    const m = text.match(/^(\d{4})[\/-](\d{1,2})$/);
    if (m) {
      const year = Number(m[1]);
      const monthIndex = Number(m[2]) - 1;
      if (Number.isFinite(year) && Number.isFinite(monthIndex) && monthIndex >= 0 && monthIndex < 12) {
        const start = new Date(year, monthIndex, 1);
        const end = new Date(year, monthIndex + 1, 0);
        return {
          key: Utilities.formatDate(start, tz, 'yyyy-MM'),
          start,
          end,
          isCurrent: start.getTime() === currentStart.getTime()
        };
      }
    }
  }
  const start = new Date(currentStart.getTime());
  const end = new Date(currentStart.getFullYear(), currentStart.getMonth() + 1, 0);
  return {
    key: Utilities.formatDate(start, tz, 'yyyy-MM'),
    start,
    end,
    isCurrent: true
  };
}

function buildVisitAttendancePayrollPortalData_(normalizedEmail, staffProfile){
  const payrollEmployeeId = staffProfile && staffProfile.payrollEmployeeId ? staffProfile.payrollEmployeeId : '';
  const payrollContext = readPayrollEmployeeRecords_();
  let employee = payrollEmployeeId ? payrollContext.mapById.get(payrollEmployeeId) : null;
  if (!employee && normalizedEmail) {
    employee = payrollContext.mapByEmail.get(normalizedEmail);
  }
  if (!employee) {
    return {
      available: false,
      restrictedReason: '給与従業員IDが未連携です。総務までお問い合わせください。'
    };
  }
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  const folder = findPayrollEmployeeFolder_(employee.name);
  let payslips = [];
  try {
    payslips = folder ? listPayrollPayslipFilesInFolder_(folder, { limit: 50, tz }) : [];
  } catch (err) {
    Logger.log('[buildVisitAttendancePayrollPortalData_] failed to list files: ' + err);
    payslips = [];
  }
  const message = payslips.length === 0 ? 'まだ給与明細が発行されていません。' : '';
  return {
    available: true,
    employee: {
      id: employee.id,
      name: employee.name,
      base: employee.base || '',
      baseKey: employee.baseKey || ''
    },
    folder: folder ? { id: folder.getId(), name: folder.getName(), url: folder.getUrl() } : null,
    payslips,
    message,
    updatedAt: formatIsoStringWithOffset_(new Date(), tz)
  };
}

function getVisitAttendancePortalData(options){
  assertDomain_();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const email = (Session.getActiveUser() || {}).getEmail() || '';
  const normalizedEmail = normalizeEmailKey_(email);
  if (!normalizedEmail) {
    throw new Error('勤怠ビューを利用するには Google アカウントでログインしてください');
  }
  const now = new Date();
  const range = resolveVisitAttendanceMonthRange_(options && options.month, tz, now);
  const staffSettings = readVisitAttendanceStaffSettings_();
  const staffProfile = resolveVisitAttendanceStaffProfile_(normalizedEmail, { staffSettings });
  const attendance = readVisitAttendanceRecordsForEmail_(normalizedEmail, { startDate: range.start, endDate: range.end, tz });
  const requests = readVisitAttendanceRequests_({ email: normalizedEmail, startDate: range.start, endDate: range.end, tz });
  const requestMap = new Map();
  requests.forEach(req => {
    if (!requestMap.has(req.targetDate)) {
      requestMap.set(req.targetDate, req);
    }
  });
  attendance.forEach(record => {
    const req = requestMap.get(record.date);
    if (req) {
      record.request = req;
    }
  });
  attendance.forEach(record => {
    if (!record) return;
    const hasWeekday = record.weekday && String(record.weekday).trim();
    if (hasWeekday) return;
    const resolvedDate = record.date ? createDateFromKey_(record.date) : null;
    if (resolvedDate) {
      record.weekday = getWeekdaySymbol_(resolvedDate, tz) || '';
    }
  });
  const totalWork = attendance.reduce((sum, r) => sum + (Number.isFinite(r.workMinutes) ? r.workMinutes : 0), 0);
  const totalBreak = attendance.reduce((sum, r) => sum + (Number.isFinite(r.breakMinutes) ? r.breakMinutes : 0), 0);
  const firstOfCurrent = new Date(now.getFullYear(), now.getMonth(), 1);
  const canRequest = range.start.getTime() < firstOfCurrent.getTime();
  const isAdmin = !!isAdminUser_();
  const pendingForAdmin = isAdmin ? readVisitAttendanceRequests_({ status: 'pending', tz }) : [];
  const adminData = isAdmin ? {
    correctionRequests: pendingForAdmin.filter(req => req.type !== VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE),
    paidLeaveRequests: pendingForAdmin.filter(req => req.type === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE)
  } : null;

  const currentYear = now.getFullYear();
  const quotaDays = resolveAnnualPaidLeaveQuota_(normalizedEmail, { staffSettings });
  const usage = calculatePaidLeaveUsageForYear_(normalizedEmail, currentYear, tz);
  const remainingDays = Math.max(0, quotaDays - (usage.usedDays || 0));
  const paidLeaveSummary = {
    year: currentYear,
    quotaDays,
    usedDays: usage.usedDays || 0,
    remainingDays,
    requiredDays: 5
  };

  let payrollPortal = null;
  try {
    payrollPortal = buildVisitAttendancePayrollPortalData_(normalizedEmail, staffProfile);
  } catch (err) {
    payrollPortal = {
      available: false,
      restrictedReason: err && err.message ? err.message : '給与明細の取得に失敗しました。'
    };
  }

  return {
    ok: true,
    user: {
      email: normalizedEmail,
      displayName: staffProfile.displayName || resolveStaffDisplayName_(normalizedEmail),
      isAdmin,
      employmentType: staffProfile.employmentType,
      employmentLabel: staffProfile.employmentLabel,
      staffProfile
    },
    timezone: tz,
    month: {
      key: range.key,
      label: Utilities.formatDate(range.start, tz, 'yyyy年M月'),
      start: Utilities.formatDate(range.start, tz, 'yyyy-MM-dd'),
      end: Utilities.formatDate(range.end, tz, 'yyyy-MM-dd'),
      isCurrent: !!range.isCurrent,
      canRequest
    },
    months: buildVisitAttendancePortalMonths_(tz, now, 12),
    attendance,
    requests,
    totals: {
      days: attendance.length,
      workMinutes: totalWork,
      workText: formatDurationText_(totalWork),
      breakMinutes: totalBreak,
      breakText: formatDurationText_(totalBreak)
    },
    policy: {
      workStart: formatMinutesAsTimeText_(VISIT_ATTENDANCE_WORK_START_MINUTES),
      workEndLimit: formatMinutesAsTimeText_(VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES),
      roundingMinutes: VISIT_ATTENDANCE_ROUNDING_MINUTES
    },
    admin: adminData,
    paidLeave: paidLeaveSummary,
    payroll: payrollPortal
  };
}

function previewPaidLeavePlan(payload){
  assertDomain_();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const email = (Session.getActiveUser() || {}).getEmail() || '';
  const normalizedEmail = normalizeEmailKey_(email);
  if (!normalizedEmail) {
    throw new Error('ログインユーザーを特定できませんでした');
  }
  const data = payload || {};
  try {
    const targetDay = parsePaidLeaveTargetDate_(data.date || data.targetDate);
    const plan = buildPaidLeavePlan_(normalizedEmail, targetDay, { tz, requestType: data.type || data.leaveType || data.kind });
    return {
      ok: true,
      plan: {
        leaveType: plan.leaveType,
        leaveLabel: plan.leaveLabel,
        workText: plan.workText,
        fullWorkText: plan.fullWorkText,
        halfWorkText: plan.halfWorkText,
        shiftText: plan.shiftText,
        employmentType: plan.employmentType,
        employmentLabel: plan.employmentLabel,
        halfAvailable: plan.halfAvailable
      },
      blockedReason: plan.blockedReason || ''
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function submitPaidLeaveRequest(payload){
  assertDomain_();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const email = (Session.getActiveUser() || {}).getEmail() || '';
  const normalizedEmail = normalizeEmailKey_(email);
  if (!normalizedEmail) {
    throw new Error('ログインユーザーを特定できませんでした');
  }

  const data = payload || {};
  const targetDateValue = data.date || data.targetDate;
  const targetDate = parsePaidLeaveTargetDate_(targetDateValue);

  const today = new Date();
  const firstOfCurrent = new Date(today.getFullYear(), today.getMonth(), 1);
  const targetDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
  if (targetDay.getTime() < firstOfCurrent.getTime()) {
    throw new Error('有給申請は当月以降の日付のみ指定できます');
  }

  const pendingOrExisting = readVisitAttendanceRequests_({ email: normalizedEmail, startDate: targetDay, endDate: targetDay, tz });
  const duplicate = pendingOrExisting.some(req => req.type === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE && req.status !== 'rejected');
  if (duplicate) {
    throw new Error('同じ日の有給申請が既に登録されています');
  }

  const existingAttendance = readVisitAttendanceRecordsForEmail_(normalizedEmail, { startDate: targetDay, endDate: targetDay, tz });
  const hasPaidLeaveRecord = existingAttendance.some(record => (record.leaveType || '').toLowerCase() === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE);
  if (hasPaidLeaveRecord) {
    throw new Error('この日は既に有給として登録されています');
  }

  const note = String(data.note || data.reason || '').trim();
  const leaveType = normalizePaidLeaveRequestType_(data.leaveType || data.type || data.kind);
  const plan = buildPaidLeavePlan_(normalizedEmail, targetDay, { tz, requestType: leaveType });
  if (plan.blockedReason) {
    throw new Error(plan.blockedReason);
  }

  const sheet = ensureVisitAttendanceRequestSheet_();
  const row = [
    Utilities.getUuid(),
    new Date(),
    normalizedEmail,
    normalizedEmail,
    targetDay,
    '有給',
    '有給',
    0,
    note,
    'pending',
    '',
    '',
    '',
    '',
    VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE
  ];
  row[13] = JSON.stringify(plan);
  sheet.appendRow(row);

  return { ok: true };
}

function submitVisitAttendanceRequest(payload){
  assertDomain_();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const email = (Session.getActiveUser() || {}).getEmail() || '';
  const normalizedEmail = normalizeEmailKey_(email);
  if (!normalizedEmail) {
    throw new Error('ログインユーザーを特定できませんでした');
  }
  const data = payload || {};
  let targetDate = data.targetDate;
  if (targetDate instanceof Date && !isNaN(targetDate.getTime())) {
    targetDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
  } else {
    const raw = String(targetDate || data.date || '').trim();
    if (!raw) {
      throw new Error('対象日を指定してください');
    }
    const normalized = raw.replace(/[\.\/]/g, '-');
    const m = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (!m) {
      throw new Error('対象日の形式が不正です (YYYY-MM-DD)');
    }
    const year = Number(m[1]);
    const monthIndex = Number(m[2]) - 1;
    const day = Number(m[3]);
    targetDate = new Date(year, monthIndex, day);
  }
  if (!(targetDate instanceof Date) || isNaN(targetDate.getTime())) {
    throw new Error('対象日の解析に失敗しました');
  }
  const today = new Date();
  const firstOfCurrent = new Date(today.getFullYear(), today.getMonth(), 1);
  const targetDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
  if (targetDay.getTime() >= firstOfCurrent.getTime()) {
    throw new Error('当月分の勤怠は修正申請できません（前月分まで）');
  }

  const startMinutes = VISIT_ATTENDANCE_WORK_START_MINUTES;
  let endMinutes = parseTimeTextToMinutes_(data.endTime != null ? data.endTime : data.end);
  if (!Number.isFinite(endMinutes)) {
    endMinutes = parseTimeTextToMinutes_(data.endMinutes);
  }
  if (!Number.isFinite(endMinutes)) {
    throw new Error('退勤時刻を HH:MM 形式で入力してください');
  }
  if (endMinutes <= startMinutes) {
    throw new Error('退勤時刻は出勤以降で指定してください');
  }
  if (endMinutes > VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES) {
    throw new Error('退勤は18:00までにしてください');
  }
  if (endMinutes % VISIT_ATTENDANCE_ROUNDING_MINUTES !== 0) {
    throw new Error('退勤時刻は15分単位で入力してください');
  }

  let breakMinutes = parseTimeTextToMinutes_(data.breakMinutes != null ? data.breakMinutes : data.break);
  if (!Number.isFinite(breakMinutes)) {
    breakMinutes = parseTimeTextToMinutes_(data.restMinutes != null ? data.restMinutes : data.rest);
  }
  if (!Number.isFinite(breakMinutes) || breakMinutes < 0) {
    breakMinutes = 0;
  }
  if (breakMinutes % VISIT_ATTENDANCE_ROUNDING_MINUTES !== 0) {
    throw new Error('休憩時間は15分単位で入力してください');
  }
  if (breakMinutes > endMinutes - startMinutes) {
    throw new Error('休憩時間が長すぎます');
  }

  const note = String(data.note || data.reason || '').trim();
  if (!note) {
    throw new Error('申請理由を入力してください');
  }

  const pending = readVisitAttendanceRequests_({ email: normalizedEmail, startDate: targetDay, endDate: targetDay, status: 'pending', tz });
  if (pending.length) {
    throw new Error('同じ日の申請が既に登録されています。管理者の対応をお待ちください。');
  }

  const original = readVisitAttendanceRecordsForEmail_(normalizedEmail, { startDate: targetDay, endDate: targetDay, tz });
  const sheet = ensureVisitAttendanceRequestSheet_();
  const row = [
    Utilities.getUuid(),
    new Date(),
    normalizedEmail,
    normalizedEmail,
    targetDay,
    formatMinutesAsTimeText_(startMinutes),
    formatMinutesAsTimeText_(endMinutes),
    breakMinutes,
    note,
    'pending',
    '',
    '',
    '',
    original && original.length ? JSON.stringify(original[0]) : '',
    VISIT_ATTENDANCE_REQUEST_TYPE_CORRECTION
  ];
  sheet.appendRow(row);

  return { ok: true };
}

function createVisitAttendanceRecord(payload){
  assertDomain_();
  if (!isAdminUser_()) {
    throw new Error('管理者権限が必要です');
  }
  const data = payload || {};
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';

  const normalizedEmail = normalizeEmailKey_(data.email || data.targetEmail || data.userEmail);
  if (!normalizedEmail) {
    throw new Error('スタッフのメールアドレスを指定してください');
  }

  let dateValue = data.date || data.targetDate;
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
    dateValue = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate());
  } else {
    const rawDate = String(dateValue || '').trim();
    if (!rawDate) {
      throw new Error('対象日を指定してください');
    }
    const normalizedDate = rawDate.replace(/[\.\/]/g, '-');
    const match = normalizedDate.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (!match) {
      throw new Error('対象日の形式が不正です (YYYY-MM-DD)');
    }
    const year = Number(match[1]);
    const monthIndex = Number(match[2]) - 1;
    const day = Number(match[3]);
    dateValue = new Date(year, monthIndex, day);
  }
  if (!(dateValue instanceof Date) || isNaN(dateValue.getTime())) {
    throw new Error('対象日の解析に失敗しました');
  }
  const targetDay = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate());
  const dateKey = Utilities.formatDate(targetDay, tz, 'yyyy-MM-dd');

  const resolveMinutes = values => {
    for (let i = 0; i < values.length; i++) {
      const minutes = parseTimeTextToMinutes_(values[i]);
      if (Number.isFinite(minutes)) {
        return minutes;
      }
    }
    return NaN;
  };

  let startMinutes = resolveMinutes([data.start, data.startTime, data.startMinutes]);
  if (!Number.isFinite(startMinutes)) {
    startMinutes = VISIT_ATTENDANCE_WORK_START_MINUTES;
  }

  let restMinutes = resolveMinutes([data.breakMinutes, data.break, data.restMinutes, data.rest]);
  if (!Number.isFinite(restMinutes) || restMinutes < 0) {
    restMinutes = 0;
  }

  let workMinutes = resolveMinutes([data.workMinutes, data.work, data.durationMinutes]);
  if (!Number.isFinite(workMinutes)) {
    const endResolved = resolveMinutes([data.end, data.endTime, data.endMinutes]);
    if (Number.isFinite(endResolved)) {
      workMinutes = Math.max(0, endResolved - startMinutes - restMinutes);
    }
  }
  if (!Number.isFinite(workMinutes) || workMinutes <= 0) {
    throw new Error('勤務時間（workMinutes）を指定してください');
  }

  let endMinutes = startMinutes + restMinutes + workMinutes;

  const rounding = VISIT_ATTENDANCE_ROUNDING_MINUTES;
  if (startMinutes % rounding !== 0 || restMinutes % rounding !== 0 || endMinutes % rounding !== 0) {
    throw new Error('時間は15分単位で指定してください');
  }
  const isHourlyStaff = toBoolean_(data.isHourlyStaff);
  const capResult = capVisitAttendanceEndMinutes_(startMinutes, restMinutes, endMinutes, { isHourlyStaff });
  if (capResult.adjusted) {
    endMinutes = capResult.endMinutes;
    if (Number.isFinite(capResult.workMinutes)) {
      workMinutes = Math.min(workMinutes, capResult.workMinutes);
    } else {
      workMinutes = Math.max(0, endMinutes - startMinutes - restMinutes);
    }
  }
  if (!isHourlyStaff) {
    endMinutes = Math.min(endMinutes, VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES);
  }
  if (endMinutes <= startMinutes) {
    throw new Error('退勤時刻は出勤以降で指定してください');
  }
  if (restMinutes > endMinutes - startMinutes) {
    throw new Error('休憩時間が長すぎます');
  }

  const breakdown = String(data.breakdown || (data.leaveType === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE ? '有給' : '') || '').trim();
  const leaveType = String(data.leaveType || '').trim();
  const isDailyStaff = toBoolean_(data.isDailyStaff);
  const sourceRaw = String(data.source || '').trim();
  let source = sourceRaw || (leaveType === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE ? VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE : 'manual');
  source = resolveVisitAttendanceRoundedSource_(source, capResult.adjusted, sourceRaw || (source === VISIT_ATTENDANCE_AUTO_FLAG_VALUE ? 'auto' : source));

  let flagValue = String(data.flag || '').trim();
  if (!flagValue) {
    if (source === VISIT_ATTENDANCE_AUTO_FLAG_VALUE || source === 'autoRounded18') {
      flagValue = VISIT_ATTENDANCE_AUTO_FLAG_VALUE;
    } else if (leaveType === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE || source === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE) {
      flagValue = VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE;
    } else {
      flagValue = 'manual';
    }
  }

  const sheet = ensureVisitAttendanceSheet_();
  const width = Math.min(VISIT_ATTENDANCE_SHEET_HEADER.length, sheet.getMaxColumns());
  const existingMap = readVisitAttendanceExistingMap_(sheet, tz);
  const key = dateKey + '||' + normalizedEmail;

  const rowValues = [
    targetDay,
    normalizedEmail,
    formatMinutesAsTimeText_(startMinutes),
    formatMinutesAsTimeText_(endMinutes),
    formatMinutesAsTimeText_(workMinutes),
    formatMinutesAsTimeText_(restMinutes),
    breakdown,
    flagValue,
    leaveType,
    isHourlyStaff,
    isDailyStaff,
    source
  ];

  let rowNumber = null;
  if (existingMap.has(key)) {
    const entry = existingMap.get(key);
    rowNumber = entry.rowNumber;
  }

  if (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, width).setValues([rowValues]);
  } else {
    rowNumber = sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, width).setValues([rowValues]);
  }

  const actor = (Session.getActiveUser() || {}).getEmail() || '';
  log_('勤怠レコード作成', normalizedEmail, JSON.stringify({ date: dateKey, leaveType, source, actor }));

  return {
    ok: true,
    date: dateKey,
    email: normalizedEmail,
    rowNumber,
    workMinutes,
    restMinutes
  };
}

function updateVisitAttendanceRecord(payload){
  assertDomain_();
  if (!isAdminUser_()) {
    throw new Error('管理者権限が必要です');
  }
  const data = payload || {};
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';

  const normalizedEmail = normalizeEmailKey_(data.email || data.targetEmail || data.userEmail);
  if (!normalizedEmail) {
    throw new Error('スタッフのメールアドレスを指定してください');
  }

  let dateValue = data.date || data.targetDate;
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
    dateValue = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate());
  } else {
    const rawDate = String(dateValue || '').trim();
    if (!rawDate) {
      throw new Error('対象日を指定してください');
    }
    const normalizedDate = rawDate.replace(/[\.\/]/g, '-');
    const match = normalizedDate.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (!match) {
      throw new Error('対象日の形式が不正です (YYYY-MM-DD)');
    }
    const year = Number(match[1]);
    const monthIndex = Number(match[2]) - 1;
    const day = Number(match[3]);
    dateValue = new Date(year, monthIndex, day);
  }
  if (!(dateValue instanceof Date) || isNaN(dateValue.getTime())) {
    throw new Error('対象日の解析に失敗しました');
  }
  const dateKey = Utilities.formatDate(dateValue, tz, 'yyyy-MM-dd');
  const targetDay = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate());

  const resolveMinutes = values => {
    for (let i = 0; i < values.length; i++) {
      const minutes = parseTimeTextToMinutes_(values[i]);
      if (Number.isFinite(minutes)) {
        return minutes;
      }
    }
    return NaN;
  };

  const startMinutes = resolveMinutes([data.start, data.startTime, data.startMinutes]);
  if (!Number.isFinite(startMinutes)) {
    throw new Error('出勤時刻を HH:MM 形式で指定してください');
  }

  let endMinutes = resolveMinutes([data.end, data.endTime, data.endMinutes]);
  if (!Number.isFinite(endMinutes)) {
    throw new Error('退勤時刻を HH:MM 形式で指定してください');
  }

  let breakMinutes = resolveMinutes([data.breakMinutes, data.break, data.restMinutes, data.rest]);
  if (!Number.isFinite(breakMinutes) || breakMinutes < 0) {
    breakMinutes = 0;
  }

  const rounding = VISIT_ATTENDANCE_ROUNDING_MINUTES;
  if (startMinutes % rounding !== 0) {
    throw new Error('出勤時刻は15分単位で指定してください');
  }
  if (endMinutes % rounding !== 0) {
    throw new Error('退勤時刻は15分単位で指定してください');
  }
  if (breakMinutes % rounding !== 0) {
    throw new Error('休憩時間は15分単位で指定してください');
  }
  if (endMinutes <= startMinutes) {
    throw new Error('退勤時刻は出勤以降で指定してください');
  }
  if (breakMinutes > endMinutes - startMinutes) {
    throw new Error('休憩時間が長すぎます');
  }

  const note = String(data.note || data.reason || '').trim();
  if (!note) {
    throw new Error('修正理由（note）を入力してください');
  }

  const sheet = ensureVisitAttendanceSheet_();
  const width = Math.min(VISIT_ATTENDANCE_SHEET_HEADER.length, sheet.getMaxColumns());
  const existingMap = readVisitAttendanceExistingMap_(sheet, tz);
  const key = dateKey + '||' + normalizedEmail;

  const resolveRow = rowNumber => {
    if (!Number.isFinite(rowNumber) || rowNumber < 2) return null;
    const range = sheet.getRange(rowNumber, 1, 1, width);
    const values = range.getValues()[0];
    const displays = range.getDisplayValues()[0];
    const rowDateKey = formatDateKeyFromValue_(values[0], tz) || formatDateKeyFromValue_(displays[0], tz);
    const rowEmail = normalizeEmailKey_(values[1] || displays[1]);
    if (rowDateKey === dateKey && rowEmail === normalizedEmail) {
      return { rowNumber, values, displays };
    }
    return null;
  };

  let targetRow = null;
  if (existingMap.has(key)) {
    targetRow = resolveRow(existingMap.get(key).rowNumber);
  }
  if (!targetRow) {
    const lastRow = sheet.getLastRow();
    for (let row = 2; row <= lastRow; row++) {
      targetRow = resolveRow(row);
      if (targetRow) break;
    }
  }
  if (!targetRow) {
    throw new Error('VisitAttendance シートの対象行を特定できませんでした');
  }

  const existingEmail = targetRow.values[1] || targetRow.displays[1] || normalizedEmail;
  const breakdownCell = targetRow.values[6] != null && targetRow.values[6] !== '' ? targetRow.values[6] : targetRow.displays[6];
  const leaveTypeCell = targetRow.values[8] != null && targetRow.values[8] !== '' ? targetRow.values[8] : (targetRow.displays[8] || '');
  const hourlyCellRaw = targetRow.values[9] != null && targetRow.values[9] !== '' ? targetRow.values[9] : targetRow.displays[9];
  const dailyCellRaw = targetRow.values[10] != null && targetRow.values[10] !== '' ? targetRow.values[10] : targetRow.displays[10];
  const sourceCellRaw = targetRow.values[11] != null && targetRow.values[11] !== '' ? targetRow.values[11] : targetRow.displays[11];
  const isHourlyStaff = toBoolean_(hourlyCellRaw);
  const capResult = capVisitAttendanceEndMinutes_(startMinutes, breakMinutes, endMinutes, { isHourlyStaff });
  if (capResult.adjusted) {
    endMinutes = capResult.endMinutes;
  }
  let workMinutes = Math.max(0, endMinutes - startMinutes - breakMinutes);
  if (capResult.adjusted && Number.isFinite(capResult.workMinutes)) {
    workMinutes = capResult.workMinutes;
  }
  if (!isHourlyStaff) {
    endMinutes = Math.min(endMinutes, VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES);
  }
  const startText = formatMinutesAsTimeText_(startMinutes);
  const endText = formatMinutesAsTimeText_(endMinutes);
  const breakText = formatMinutesAsTimeText_(breakMinutes);
  const workText = formatMinutesAsTimeText_(workMinutes);
  const hourlyCell = isHourlyStaff ? true : '';
  const dailyCell = toBoolean_(dailyCellRaw) ? true : '';
  const fallbackSource = String(sourceCellRaw || '').trim();
  const sourceCell = resolveVisitAttendanceRoundedSource_(fallbackSource, capResult.adjusted, fallbackSource || 'manual') || 'manual';

  const newRow = [
    targetDay,
    existingEmail,
    startText,
    endText,
    workText,
    breakText,
    breakdownCell,
    'manual',
    leaveTypeCell,
    hourlyCell,
    dailyCell,
    sourceCell
  ];

  sheet.getRange(targetRow.rowNumber, 1, 1, width).setValues([newRow]);

  const actor = (Session.getActiveUser() || {}).getEmail() || '';
  const logDetail = JSON.stringify({
    date: dateKey,
    email: normalizedEmail,
    start: startText,
    end: endText,
    break: breakText,
    work: workText,
    note,
    actor
  });
  log_('勤怠手動修正', normalizedEmail, logDetail);
  Logger.log('[updateVisitAttendanceRecord] ' + logDetail);

  return {
    ok: true,
    rowNumber: targetRow.rowNumber,
    date: dateKey,
    email: normalizedEmail,
    start: startText,
    end: endText,
    breakMinutes,
    workMinutes
  };
}

function updateVisitAttendanceRequestStatus(payload){
  assertDomain_();
  if (!isAdminUser_()) {
    throw new Error('管理者権限が必要です');
  }
  const data = payload || {};
  const id = String(data.id || data.requestId || '').trim();
  if (!id) {
    throw new Error('申請IDが不正です');
  }
  const statusRaw = String(data.status || '').trim().toLowerCase();
  if (!statusRaw) {
    throw new Error('状態を指定してください');
  }
  if (['pending','approved','rejected'].indexOf(statusRaw) === -1) {
    throw new Error('状態は pending / approved / rejected のいずれかです');
  }
  const note = String(data.note || '').trim();
  const sheet = ensureVisitAttendanceRequestSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('申請が見つかりません');
  }
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < ids.length; i++) {
    const value = String(ids[i][0] || '').trim();
    if (value === id) {
      targetRow = i + 2;
      break;
    }
  }
  if (targetRow === -1) {
    throw new Error('対象の申請が見つかりません');
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const width = Math.min(VISIT_ATTENDANCE_REQUEST_SHEET_HEADER.length, sheet.getMaxColumns());
  const requestRange = sheet.getRange(targetRow, 1, 1, width);
  const requestRow = requestRange.getValues()[0];
  const requestDisplay = requestRange.getDisplayValues()[0];

  const requestTypeRaw = String((requestRow[14] != null && requestRow[14] !== '') ? requestRow[14] : (requestDisplay[14] || '')).trim().toLowerCase();
  const requestType = requestTypeRaw || VISIT_ATTENDANCE_REQUEST_TYPE_CORRECTION;

  const targetEmail = normalizeEmailKey_(requestRow[3] || requestDisplay[3] || requestRow[2] || requestDisplay[2]);
  if (!targetEmail) {
    throw new Error('対象メールを特定できませんでした');
  }

  let targetDate = requestRow[4];
  if (!(targetDate instanceof Date) || isNaN(targetDate.getTime())) {
    const key = formatDateKeyFromValue_(requestRow[4], tz) || formatDateKeyFromValue_(requestDisplay[4], tz);
    targetDate = createDateFromKey_(key || '');
  }
  if (!(targetDate instanceof Date) || isNaN(targetDate.getTime())) {
    throw new Error('対象日を解析できませんでした');
  }
  const targetDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
  const dateKey = Utilities.formatDate(targetDay, tz, 'yyyy-MM-dd');

  if (statusRaw === 'approved') {
    if (requestType === VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE) {
      const planMeta = parsePaidLeaveRequestMetadata_(requestRow[13] != null && requestRow[13] !== '' ? requestRow[13] : requestDisplay[13]);
      let startMinutes = Number(planMeta && planMeta.startMinutes);
      if (!Number.isFinite(startMinutes)) {
        startMinutes = VISIT_ATTENDANCE_WORK_START_MINUTES;
      }
      let workMinutes = Number(planMeta && planMeta.workMinutes);
      if (!Number.isFinite(workMinutes) || workMinutes <= 0) {
        workMinutes = Number(data.workMinutes);
      }
      if (!Number.isFinite(workMinutes) || workMinutes <= 0) {
        workMinutes = PAID_LEAVE_DEFAULT_WORK_MINUTES;
      }
      let restMinutes = Number(planMeta && planMeta.breakMinutes);
      if (!Number.isFinite(restMinutes) || restMinutes < 0) {
        restMinutes = 0;
      }
      const employmentType = String(planMeta && planMeta.employmentType || '').toLowerCase();
      const isHourlyStaff = employmentType === 'parttime';
      const isDailyStaff = employmentType === 'daily';
      createVisitAttendanceRecord({
        email: targetEmail,
        date: targetDay,
        startMinutes,
        workMinutes,
        restMinutes,
        leaveType: VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE,
        source: VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE,
        flag: VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE,
        breakdown: '有給',
        isHourlyStaff,
        isDailyStaff
      });
    } else {
      const attendanceSheet = ensureVisitAttendanceSheet_();
      const attendanceWidth = Math.min(VISIT_ATTENDANCE_SHEET_HEADER.length, attendanceSheet.getMaxColumns());

      let endMinutes = parseTimeTextToMinutes_(requestRow[6] != null && requestRow[6] !== '' ? requestRow[6] : requestDisplay[6]);
      if (!Number.isFinite(endMinutes)) {
        endMinutes = parseTimeTextToMinutes_(payload.endMinutes != null ? payload.endMinutes : payload.end);
      }
      if (!Number.isFinite(endMinutes)) {
        throw new Error('退勤時刻を解析できませんでした');
      }
      if (endMinutes % VISIT_ATTENDANCE_ROUNDING_MINUTES !== 0) {
        throw new Error('退勤時刻は15分単位である必要があります');
      }
      if (endMinutes > VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES) {
        throw new Error('退勤時刻が制限を超えています');
      }

      let breakMinutes = parseTimeTextToMinutes_(requestRow[7] != null && requestRow[7] !== '' ? requestRow[7] : requestDisplay[7]);
      if (!Number.isFinite(breakMinutes) || breakMinutes < 0) {
        breakMinutes = 0;
      }
      if (breakMinutes % VISIT_ATTENDANCE_ROUNDING_MINUTES !== 0) {
        throw new Error('休憩時間は15分単位である必要があります');
      }

      const startMinutes = VISIT_ATTENDANCE_WORK_START_MINUTES;
      if (endMinutes <= startMinutes) {
        throw new Error('退勤時刻は出勤以降で指定してください');
      }
      if (breakMinutes > endMinutes - startMinutes) {
        throw new Error('休憩時間が長すぎます');
      }

      let workMinutes = Math.max(0, endMinutes - startMinutes - breakMinutes);

      let originalData = null;
      const originalRaw = requestRow[13] != null && requestRow[13] !== '' ? requestRow[13] : requestDisplay[13];
      if (originalRaw != null && originalRaw !== '') {
        if (typeof originalRaw === 'string') {
          try {
            originalData = JSON.parse(originalRaw);
          } catch (err) {
            originalData = null;
          }
        } else {
          originalData = originalRaw;
        }
      }

      const resolveAttendanceRow = rowNumber => {
        if (!Number.isFinite(rowNumber) || rowNumber < 2) return null;
        const range = attendanceSheet.getRange(rowNumber, 1, 1, attendanceWidth);
        const values = range.getValues()[0];
        const displays = range.getDisplayValues()[0];
        const rowDateKey = formatDateKeyFromValue_(values[0], tz) || formatDateKeyFromValue_(displays[0], tz);
        const rowEmail = normalizeEmailKey_(values[1] || displays[1]);
        if (rowDateKey === dateKey && rowEmail === targetEmail) {
          return { rowNumber, values, displays };
        }
        return null;
      };

      let attendanceRow = null;
      if (originalData && typeof originalData === 'object') {
        const candidate = Number(originalData.rowNumber || originalData.row || originalData.rowIndex);
        attendanceRow = resolveAttendanceRow(candidate);
      }
      if (!attendanceRow) {
        const existingMap = readVisitAttendanceExistingMap_(attendanceSheet, tz);
        const entry = existingMap.get(dateKey + '||' + targetEmail);
        if (entry) {
          attendanceRow = resolveAttendanceRow(entry.rowNumber);
        }
      }
      if (!attendanceRow) {
        throw new Error('VisitAttendance シートの対象行を特定できませんでした');
      }

      const emailCell = attendanceRow.values[1] || attendanceRow.displays[1] || requestRow[3] || requestDisplay[3] || requestRow[2] || requestDisplay[2] || '';
      const breakdownCell = attendanceRow.values[6] != null && attendanceRow.values[6] !== '' ? attendanceRow.values[6] : attendanceRow.displays[6];
      const flagCell = attendanceRow.values[7] != null && attendanceRow.values[7] !== '' ? attendanceRow.values[7] : attendanceRow.displays[7];
      const leaveTypeCell = attendanceRow.values[8] != null && attendanceRow.values[8] !== '' ? attendanceRow.values[8] : (attendanceRow.displays[8] || '');
      const hourlyCellRaw = attendanceRow.values[9] != null && attendanceRow.values[9] !== '' ? attendanceRow.values[9] : attendanceRow.displays[9];
      const dailyCellRaw = attendanceRow.values[10] != null && attendanceRow.values[10] !== '' ? attendanceRow.values[10] : attendanceRow.displays[10];
      const sourceCellRaw = attendanceRow.values[11] != null && attendanceRow.values[11] !== '' ? attendanceRow.values[11] : attendanceRow.displays[11];
      const isHourlyStaff = toBoolean_(hourlyCellRaw);
      const capResult = capVisitAttendanceEndMinutes_(startMinutes, breakMinutes, endMinutes, { isHourlyStaff });
      if (capResult.adjusted) {
        endMinutes = capResult.endMinutes;
      }
      if (capResult.adjusted && Number.isFinite(capResult.workMinutes)) {
        workMinutes = capResult.workMinutes;
      } else {
        workMinutes = Math.max(0, endMinutes - startMinutes - breakMinutes);
      }
      if (!isHourlyStaff) {
        endMinutes = Math.min(endMinutes, VISIT_ATTENDANCE_WORK_END_LIMIT_MINUTES);
      }
      const hourlyCell = isHourlyStaff ? true : '';
      const dailyCell = toBoolean_(dailyCellRaw) ? true : '';
      const fallbackSource = String(sourceCellRaw || '').trim() || (String(flagCell || '').trim().toLowerCase() === VISIT_ATTENDANCE_AUTO_FLAG_VALUE ? 'auto' : 'manual');
      const sourceCell = resolveVisitAttendanceRoundedSource_(fallbackSource, capResult.adjusted, fallbackSource) || fallbackSource || 'manual';

      const newRowValues = [
        targetDay,
        emailCell,
        formatMinutesAsTimeText_(startMinutes),
        formatMinutesAsTimeText_(endMinutes),
        formatMinutesAsTimeText_(workMinutes),
        formatMinutesAsTimeText_(breakMinutes),
        breakdownCell,
        flagCell,
        leaveTypeCell,
        hourlyCell,
        dailyCell,
        sourceCell
      ];

      attendanceSheet.getRange(attendanceRow.rowNumber, 1, 1, attendanceWidth).setValues([newRowValues]);
    }
  }

  const now = new Date();
  const actor = (Session.getActiveUser() || {}).getEmail() || '';
  sheet.getRange(targetRow, 10).setValue(statusRaw);
  sheet.getRange(targetRow, 11).setValue(now);
  sheet.getRange(targetRow, 12).setValue(actor);
  sheet.getRange(targetRow, 13).setValue(note);
  return { ok: true };
}

function approvePaidLeaveRequest(payload){
  assertDomain_();
  if (!isAdminUser_()) {
    throw new Error('管理者権限が必要です');
  }
  const data = payload || {};
  const id = String(data.id || data.requestId || '').trim();
  if (!id) {
    throw new Error('申請IDが不正です');
  }
  const workMinutes = Number(data.workMinutes);
  const note = String(data.note || '').trim();
  return updateVisitAttendanceRequestStatus({ id, status: 'approved', note, workMinutes });
}

function rejectPaidLeaveRequest(payload){
  assertDomain_();
  if (!isAdminUser_()) {
    throw new Error('管理者権限が必要です');
  }
  const data = payload || {};
  const id = String(data.id || data.requestId || '').trim();
  if (!id) {
    throw new Error('申請IDが不正です');
  }
  const note = String(data.note || '').trim();
  return updateVisitAttendanceRequestStatus({ id, status: 'rejected', note });
}

function runVisitAttendanceSyncJob(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('[runVisitAttendanceSyncJob] ロック取得に失敗しました');
    return null;
  }
  try {
    const summary = syncVisitAttendance();
    Logger.log('[runVisitAttendanceSyncJob] ' + JSON.stringify(summary));
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function ensureVisitAttendanceSyncTrigger(){
  const handler = 'runVisitAttendanceSyncJob';
  const triggers = ScriptApp.getProjectTriggers();
  let hasClockTrigger = false;
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === handler) {
      if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
        hasClockTrigger = true;
      } else {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  });
  if (!hasClockTrigger) {
    ScriptApp.newTrigger(handler)
      .timeBased()
      .everyDays(1)
      .atHour(0)
      .create();
    Logger.log('[ensureVisitAttendanceSyncTrigger] 新規トリガーを作成しました (00:00 JST)');
  }
  return true;
}

function formatVisitAttendanceSyncSummary_(summary){
  if (!summary) {
    return '勤怠データの同期は実行されませんでした。';
  }
  const lines = [
    '勤怠データの同期を実行しました。',
    '対象行: ' + (summary.targetedRows || 0),
    '新規追加: ' + (summary.appended || 0),
    '更新: ' + (summary.updated || 0),
    '手動調整のためスキップ: ' + (summary.manualSkipped || 0),
    'エラー: ' + (summary.errors || 0)
  ];
  return lines.join('\n');
}

function runVisitAttendanceSyncJobFromMenu(){
  const ui = SpreadsheetApp.getUi();
  try {
    const summary = runVisitAttendanceSyncJob();
    if (summary === null) {
      ui.alert('別の同期処理が実行中のため、今回の実行はスキップされました。');
      return;
    }
    ui.alert(formatVisitAttendanceSyncSummary_(summary));
  } catch (err) {
    const message = err && err.message ? err.message : String(err);
    Logger.log('[runVisitAttendanceSyncJobFromMenu] ' + message);
    ui.alert('勤怠同期でエラーが発生しました: ' + message);
  }
}

function ensureVisitAttendanceSyncTriggerFromMenu(){
  const ui = SpreadsheetApp.getUi();
  try {
    ensureVisitAttendanceSyncTrigger();
    ui.alert('日次トリガーを確認しました（未設定の場合は新規作成されました）。');
  } catch (err) {
    const message = err && err.message ? err.message : String(err);
    Logger.log('[ensureVisitAttendanceSyncTriggerFromMenu] ' + message);
    ui.alert('トリガーの確認に失敗しました: ' + message);
  }
}

function normalizeTreatmentTimestamp_(value, tz) {
  if (value instanceof Date) {
    return value;
  }
  const str = String(value || '').trim();
  if (!str) return null;
  const iso = str.replace(' ', 'T');
  const date = new Date(iso + (iso.endsWith('Z') || iso.includes('+') ? '' : (tz === 'Asia/Tokyo' ? '+09:00' : 'Z')));
  if (!isNaN(date.getTime())) {
    return date;
  }
  try {
    return new Date(str);
  } catch (e) {
    return null;
  }
}

function findTreatmentRowById_(sheet, treatmentId) {
  if (!treatmentId) return null;
  const lr = sheet.getLastRow();
  if (lr < 2) return null;
  const ids = sheet.getRange(2, 7, lr - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const id = String(ids[i][0] || '').trim();
    if (id === treatmentId) {
      const rowNumber = i + 2;
      const row = sheet.getRange(rowNumber, 1, 1, 7).getValues()[0];
      return { rowNumber, row };
    }
  }
  return null;
}

/***** 申し送り：内部ユーティリティ *****/
// 申し送りタブを安全に取得（無ければ作成＋ヘッダ付与）
function ensureHandoverSheet_(){
  const wb = ss();                                  // ← 既存の ss() を使用
  let s = wb.getSheetByName('申し送り');
  if (!s) s = wb.insertSheet('申し送り');
  if (s.getLastRow() === 0) {
    s.getRange(1,1,1,5).setValues([['TS','患者ID','ユーザー','メモ','FileIds']]);
  }
  return s;
}

// 画像保存ルートフォルダを解決
// 優先: ScriptProperty(HANDOVER_FOLDER_ID) → APP.PARENT_FOLDER_ID → スプレッドシートと同じ親フォルダ
function getHandoverRootFolder_(){
  const propId = (PropertiesService.getScriptProperties().getProperty('HANDOVER_FOLDER_ID') || '').trim();
  try { if (propId) return DriveApp.getFolderById(propId); } catch(e){}
  try { if (APP.PARENT_FOLDER_ID) return DriveApp.getFolderById(APP.PARENT_FOLDER_ID); } catch(e){}
  return getParentFolder_();                        // ← 既存の親フォルダ解決関数を流用
}

/***** 申し送り：保存 *****/
function saveHandover(payload) {
  const s = ensureHandoverSheet_();

  const pid = String(payload && payload.patientId || '').trim();
  if (!pid) throw new Error('patientIdが空です');

  const user = (Session.getActiveUser()||{}).getEmail() || '';
  const tz   = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const now  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  const files = Array.isArray(payload && payload.files) ? payload.files : [];
  const fileIds = [];

  if (files.length){
    // ルート/申し送り/patientId の順にフォルダを用意
    const root = getHandoverRootFolder_();
    const itH = root.getFoldersByName('申し送り');
    const handoverRoot = itH.hasNext() ? itH.next() : root.createFolder('申し送り');

    const itP = handoverRoot.getFoldersByName(pid);
    const patientFolder = itP.hasNext() ? itP.next() : handoverRoot.createFolder(pid);

    files.forEach(f=>{
      try{
        // dataURL or base64 どちらでもOKにする
        const raw = String(f.data || '');
        const b64 = raw.indexOf(',') >= 0 ? raw.split(',')[1] : raw;
        const name = (f.name || 'upload.jpg');
        const blob = Utilities.newBlob(
          Utilities.base64Decode(b64),
          (f.type || 'application/octet-stream'),
          name
        );
        const saved = patientFolder.createFile(blob)
        .setName(now.replace(/[^\d]/g,'') + '_' + name);
        saved.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileIds.push(saved.getId());

      }catch(e){
        Logger.log('[handover upload error] ' + e);
      }
    });
  }

  s.appendRow([ now, pid, user, String(payload && payload.note || ''), fileIds.join(',') ]);

  const monthKey = now.slice(0, 7);
  let cleared = 0;
  try {
    const clearedMonthly = clearMonthlyHandoverReminder_(pid, monthKey);
    cleared += clearedMonthly;
    if (!clearedMonthly) {
      cleared += clearMonthlyHandoverReminder_(pid);
    }
  } catch (e) {
    Logger.log('[saveHandover] failed to clear monthly reminder: ' + (e && e.message ? e.message : e));
  }
  try {
    cleared += clearDoctorReportMissingReminder_(pid);
  } catch (e) {
    Logger.log('[saveHandover] failed to clear doctor report missing reminder: ' + (e && e.message ? e.message : e));
  }
  try {
    cleared += markNewsClearedByType(pid, '同意', { metaType: 'handover' });
  } catch (e) {
    Logger.log('[saveHandover] failed to clear handover consent news: ' + (e && e.message ? e.message : e));
  }

  return { ok:true, fileIds, cleared };
}
/***** 申し送り：一覧取得 *****/
function listHandovers(pid) {
  const s = ensureHandoverSheet_();
  const lr = s.getLastRow();
  if (lr < 2) return [];

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const vals = s.getRange(2, 1, lr - 1, 5).getValues(); // [TS, 患者ID, ユーザー, メモ, FileIds]

  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const row = i + 2; // 2行目から始まるので +2
    const [ts, id, user, note, fileIdsStr] = vals[i];
    if (String(id) !== String(pid)) continue;

    const when = ts instanceof Date
      ? Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm')
      : String(ts || '');

    const fileIds = String(fileIdsStr || '').split(',').filter(Boolean);
    const files = fileIds.map(fid => {
      try {
        const f = DriveApp.getFileById(fid);
        return "https://drive.google.com/thumbnail?id=" + f.getId() + "&sz=w300";
      } catch (e) {
        return null;
     }
    }).filter(Boolean);


    out.push({ row, when, user, note, files });
  }
  return out.reverse(); // 新しい順
}

function updateHandover(row, newNote) {
  const s = ensureHandoverSheet_();
  if (row <= 1 || row > s.getLastRow()) throw new Error('行が不正です');
  s.getRange(row, 4).setValue(newNote); // 4列目=メモ
  return true;
}
function deleteHandover(row) {
  const s = ensureHandoverSheet_();
  if (row <= 1 || row > s.getLastRow()) throw new Error('行が不正です');
  s.deleteRow(row);
  return true;
}
function seedAlbyteTestDataForQa(){
  return wrapAlbyteResponse_('albyteSeedTestData', () => {
    return withAlbyteLock_(() => {
      const tz = getConfig('timezone') || 'Asia/Tokyo';
      const now = new Date();
      const staffSpecs = [
        { name: 'テストA', pin: '1111', staffType: 'hourly', shiftEnd: '16:15' },
        { name: 'テストB', pin: '2222', staffType: 'hourly', shiftEnd: '13:30' },
        { name: 'テストC', pin: '3333', staffType: 'hourly', shiftEnd: '16:00' }
      ];

      const staffResult = ensureAlbyteTestStaff_(staffSpecs, now);
      const shiftResult = seedAlbyteTestShifts_(staffResult.records, tz, now);
      const attendanceResult = seedAlbyteTestAttendance_(staffResult.records, tz, now);

      return {
        ok: true,
        message: 'アルバイト勤怠テストデータを生成しました。',
        staffCreated: staffResult.created,
        shiftsCreated: shiftResult.created,
        attendanceCreated: attendanceResult.created
      };
    });
  });
}

function ensureAlbyteTestStaff_(specs, now){
  const context = readAlbyteStaffRecords_();
  const sheet = context.sheet;
  let created = 0;
  specs.forEach(spec => {
    const normalized = normalizeAlbyteName_(spec.name);
    if (normalized && context.mapByName.get(normalized)) {
      return;
    }
    sheet.appendRow([
      Utilities.getUuid(),
      spec.name,
      spec.pin,
      '',
      0,
      '',
      now,
      spec.staffType,
      spec.shiftEnd
    ]);
    created++;
  });
  const refreshed = readAlbyteStaffRecords_();
  const records = specs.map(spec => {
    const normalized = normalizeAlbyteName_(spec.name);
    const record = normalized ? refreshed.mapByName.get(normalized) : null;
    if (!record) {
      throw new Error('スタッフ "' + spec.name + '" の取得に失敗しました。');
    }
    return record;
  });
  return { records, created };
}

function seedAlbyteTestShifts_(staffRecords, tz, now){
  if (!staffRecords || !staffRecords.length) {
    return { created: 0 };
  }
  const context = readAlbyteShiftRecords_();
  const sheet = context.sheet;
  const width = ALBYTE_SHIFT_SHEET_HEADER.length;
  const existing = new Map();
  context.records.forEach(record => {
    if (record && record.staffId && record.dateKey) {
      existing.set(record.staffId + '::' + record.dateKey, true);
    }
  });
  const monthDays = {
    10: [3, 7, 14],
    11: [5, 12, 19]
  };
  const startOptions = ['09:00', '10:30', '11:30'];
  const shiftRows = [];
  staffRecords.forEach(staff => {
    Object.keys(monthDays).forEach(monthKey => {
      const days = monthDays[monthKey];
      days.forEach(day => {
        const dateKey = Utilities.formatDate(new Date(2025, Number(monthKey) - 1, day), tz, 'yyyy-MM-dd');
        const mapKey = staff.id + '::' + dateKey;
        if (existing.has(mapKey)) {
          return;
        }
        existing.set(mapKey, true);
        const start = startOptions[Math.floor(Math.random() * startOptions.length)];
        const end = staff.shiftEndTime && staff.shiftEndTime.trim()
          ? staff.shiftEndTime
          : (start === '09:00' ? '16:15' : '15:00');
        shiftRows.push([
          Utilities.getUuid(),
          dateKey,
          staff.id,
          staff.name,
          start,
          end,
          'テストシフト',
          now
        ]);
      });
    });
  });
  if (!shiftRows.length) {
    return { created: 0 };
  }
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, shiftRows.length, width).setValues(shiftRows);
  return { created: shiftRows.length };
}

function seedAlbyteTestAttendance_(staffRecords, tz, now){
  if (!staffRecords || !staffRecords.length) {
    return { created: 0 };
  }
  const attendanceContext = readAlbyteAttendanceRecords_({ fromDateKey: '2025-10-01', toDateKey: '2025-11-30' });
  const sheet = attendanceContext.sheet;
  const width = ALBYTE_ATTENDANCE_SHEET_HEADER.length;
  const existing = new Map();
  attendanceContext.records.forEach(record => {
    if (record && record.staffId && record.date) {
      existing.set(record.staffId + '::' + record.date, true);
    }
  });
  const staffMap = new Map();
  staffRecords.forEach(staff => {
    staffMap.set(staff.id, staff);
  });
  const monthDays = {
    10: [2, 6, 10, 17, 24],
    11: [3, 7, 13, 20, 27]
  };
  const startOptions = ['09:00', '10:30', '11:30'];
  const endOptions = {
    '09:00': ['13:30', '15:00', '16:15'],
    '10:30': ['15:00', '16:15'],
    '11:30': ['15:00', '16:15']
  };
  const breakOptions = [30, 45, 60];
  const scenarioQueue = [
    { type: 'overtime', note: '18:00超過調整テスト' },
    { type: 'missingBreak', note: '休憩未登録テスト' },
    { type: 'missingClockOut', note: '退勤未登録テスト' },
    { type: 'invalidClockOut', note: '退勤時刻バリデーションテスト' }
  ];
  const attendanceRows = [];
  const insertedKeys = [];

  staffRecords.forEach(staff => {
    Object.keys(monthDays).forEach(monthKey => {
      monthDays[monthKey].forEach(day => {
        const dateKey = Utilities.formatDate(new Date(2025, Number(monthKey) - 1, day), tz, 'yyyy-MM-dd');
        const mapKey = staff.id + '::' + dateKey;
        if (existing.has(mapKey)) {
          return;
        }
        existing.set(mapKey, true);
        const scenario = scenarioQueue.length ? scenarioQueue.shift() : null;
        const start = startOptions[Math.floor(Math.random() * startOptions.length)];
        const endPool = endOptions[start] || ['15:00'];
        let end = endPool[Math.floor(Math.random() * endPool.length)];
        let breakMinutes = breakOptions[Math.floor(Math.random() * breakOptions.length)];
        let note = 'テストデータ';
        let includeBreakLog = true;
        if (scenario) {
          note = scenario.note;
          if (scenario.type === 'overtime') {
            end = '18:30';
          } else if (scenario.type === 'missingBreak') {
            breakMinutes = '';
            includeBreakLog = false;
          } else if (scenario.type === 'missingClockOut') {
            end = '';
          } else if (scenario.type === 'invalidClockOut') {
            end = '07:45';
          }
        }
        const breakValue = typeof breakMinutes === 'number' ? breakMinutes : '';
        const logEntries = [];
        const clockInIso = buildIsoForAlbyteSeed_(dateKey, start, tz);
        if (clockInIso) {
          logEntries.push({ type: 'clockIn', at: clockInIso });
        }
        if (includeBreakLog && typeof breakMinutes === 'number') {
          const breakIso = buildIsoForAlbyteSeed_(dateKey, addMinutesToTimeText_(start, 90), tz);
          logEntries.push({ type: 'breakUpdate', at: breakIso, minutes: breakMinutes, source: 'seed' });
        }
        if (end) {
          const clockOutIso = buildIsoForAlbyteSeed_(dateKey, end, tz);
          if (clockOutIso) {
            logEntries.push({ type: 'clockOut', at: clockOutIso });
          }
        }
        attendanceRows.push([
          Utilities.getUuid(),
          staff.id,
          staff.name,
          dateKey,
          start,
          end,
          breakValue,
          note,
          '',
          serializeAlbyteAttendanceLog_(logEntries),
          now,
          now
        ]);
        insertedKeys.push({ staffId: staff.id, dateKey });
      });
    });
  });
  if (!attendanceRows.length) {
    return { created: 0 };
  }
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, attendanceRows.length, width).setValues(attendanceRows);
  const shiftContext = readAlbyteShiftRecords_();
  insertedKeys.forEach(entry => {
    const record = readAlbyteAttendanceRowFor_(entry.staffId, entry.dateKey, { sheet });
    if (record) {
      applyAlbyteAutoAdjustmentsForRow_(record, {
        sheet,
        staff: staffMap.get(entry.staffId),
        shiftContext
      });
    }
  });
  return { created: attendanceRows.length };
}

function buildIsoForAlbyteSeed_(dateKey, timeText, tz){
  if (!dateKey || !timeText) {
    return '';
  }
  const [y, m, d] = dateKey.split('-').map(Number);
  const [hh, mm] = timeText.split(':').map(Number);
  if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) {
    return '';
  }
  const hour = Number.isFinite(hh) ? hh : 0;
  const minute = Number.isFinite(mm) ? mm : 0;
  const date = new Date(y, m - 1, d, hour, minute);
  return formatIsoStringWithOffset_(date, tz || getConfig('timezone') || 'Asia/Tokyo');
}

function addMinutesToTimeText_(timeText, minutes){
  const [hh, mm] = String(timeText || '00:00').split(':').map(Number);
  if (!Number.isFinite(minutes)) {
    return timeText;
  }
  const base = new Date(2025, 0, 1, Number.isFinite(hh) ? hh : 0, Number.isFinite(mm) ? mm : 0, 0, 0);
  base.setMinutes(base.getMinutes() + minutes);
  const hour = base.getHours();
  const minute = base.getMinutes();
  const pad = value => String(value).padStart(2, '0');
  return pad(hour) + ':' + pad(minute);
}
