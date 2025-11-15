/***** ── 設定 ─────────────────────────────────*****/
const APP = {
  // Driveに保存するPDFの親フォルダID（空でも可：スプレッドシートと同じ階層に保存）
  PARENT_FOLDER_ID: '1VAv9ZOLB7A__m8ErFDPhFHvhpO21OFPP',
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
  '',
  '【構成ルール】',
  '・出力は必ず次の3つの見出しで構成してください（見出し行もそのまま出力します）。',
  '  ■施術の内容・頻度',
  '  ■患者の状態・経過',
  '  ■特記すべき事項',
  '',
  '【内容ルール】',
  '・「施術の内容・頻度」では、冒頭で必ず「頂いている『〇〇』の同意に対して、〜」という形で同意内容を引用し、施術目的と関連付けて記載してください。',
  '・「施術の内容・頻度」では、同意内容 → 施術目的（可動域改善・筋力強化・疼痛緩和など） → 頻度（週2回など）の順に簡潔に述べてください。',
  '・同意内容が空欄の場合は、「同意内容の記載なし」とせず、施術目的と頻度のみで自然に構成してください。',
  '・「患者の状態・経過」では観察事実を中心に、改善傾向・課題・留意点を簡潔に述べてください。',
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

const AUX_SHEETS_INIT_KEY = 'AUX_SHEETS_INIT_V202502';
const PATIENT_CACHE_TTL_SECONDS = 90;
const PATIENT_CACHE_KEYS = {
  header: pid => 'patient:header:' + normId_(pid),
  news: pid => 'patient:news:' + normId_(pid),
  treatments: pid => 'patient:treatments:' + normId_(pid),
  reports: pid => 'patient:reports:' + normId_(pid),
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
  '勤怠反映フラグ'
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
const VISIT_ATTENDANCE_STAFF_SHEET_HEADER = ['メール','表示名','年間有給付与日数'];
const DEFAULT_ANNUAL_PAID_LEAVE_DAYS = 10;
const PAID_LEAVE_DEFAULT_WORK_MINUTES = 8 * 60;

const ALBYTE_STAFF_SHEET_NAME = 'AlbyteStaff';
const ALBYTE_ATTENDANCE_SHEET_NAME = 'AlbyteAttendance';
const ALBYTE_STAFF_SHEET_HEADER = ['ID','名前','PIN','ロック中','連続失敗','最終ログイン','更新TS'];
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
const ALBYTE_MAX_PIN_ATTEMPTS = 5;
const ALBYTE_SESSION_SECRET_PROPERTY_KEY = 'ALBYTE_SESSION_SECRET';
const ALBYTE_SESSION_TTL_MILLIS = 1000 * 60 * 60 * 12;
const ALBYTE_BREAK_MINUTES_PRESETS = Object.freeze([30, 45, 60, 90, 120, 180]);
const ALBYTE_BREAK_STEP_MINUTES = 15;
const ALBYTE_MAX_BREAK_MINUTES = 180;

const ALBYTE_STAFF_COLUMNS = Object.freeze({
  id: 0,
  name: 1,
  pin: 2,
  locked: 3,
  failCount: 4,
  lastLogin: 5,
  updatedAt: 6
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
  return keys;
}

function invalidateGlobalNewsCache_(){
  invalidateCacheKeys_([GLOBAL_NEWS_CACHE_KEY]);
}

/***** 先頭行（見出し）の揺れに耐えるためのラベル候補群 *****/
const LABELS = {
  recNo:     ['施術録番号','施術録No','施術録NO','記録番号','カルテ番号','患者ID','患者番号'],
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
    const need = ['施術録','患者情報','News','フラグ','予定','操作ログ','定型文','添付索引','年次確認','ダッシュボード','AI報告書', VISIT_ATTENDANCE_SHEET_NAME, ALBYTE_ATTENDANCE_SHEET_NAME, ALBYTE_STAFF_SHEET_NAME];
    need.forEach(n => { if (!wb.getSheetByName(n)) wb.insertSheet(n); });

    const ensureHeader = (name, header) => {
      const s = wb.getSheetByName(name);
      if (s.getLastRow() === 0) s.appendRow(header);
    };

    // 既存タブ
    ensureHeader('施術録',   TREATMENT_SHEET_HEADER);
    ensureHeader('News',     ['TS','患者ID','種別','メッセージ','cleared','meta']);

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
    upgradeHeader('News',   ['TS','患者ID','種別','メッセージ','cleared','meta']);
    upgradeHeader('AI報告書', AI_REPORT_SHEET_HEADER);
    upgradeHeader(VISIT_ATTENDANCE_SHEET_NAME, VISIT_ATTENDANCE_SHEET_HEADER);
    upgradeHeader(ALBYTE_ATTENDANCE_SHEET_NAME, ALBYTE_ATTENDANCE_SHEET_HEADER);
    upgradeHeader(ALBYTE_STAFF_SHEET_NAME, ALBYTE_STAFF_SHEET_HEADER);
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

    props.setProperty(AUX_SHEETS_INIT_KEY, '1');
  } finally {
    if (locked) {
      lock.releaseLock();
    }
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

function parseDateValue_(value){
  if (value instanceof Date) return value;
  if (value == null || value === '') return null;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
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
      const record = {
        rowIndex,
        id,
        name,
        normalizedName,
        pin,
        locked,
        failCount,
        lastLogin,
        updatedAt
      };
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

function parseAlbyteAttendanceRow_(row, rowIndex){
  return {
    rowIndex,
    id: String(row[ALBYTE_ATTENDANCE_COLUMNS.id] || '').trim(),
    staffId: String(row[ALBYTE_ATTENDANCE_COLUMNS.staffId] || '').trim(),
    staffName: String(row[ALBYTE_ATTENDANCE_COLUMNS.staffName] || '').trim(),
    date: String(row[ALBYTE_ATTENDANCE_COLUMNS.date] || '').trim(),
    clockIn: String(row[ALBYTE_ATTENDANCE_COLUMNS.clockIn] || '').trim(),
    clockOut: String(row[ALBYTE_ATTENDANCE_COLUMNS.clockOut] || '').trim(),
    breakMinutes: Number(row[ALBYTE_ATTENDANCE_COLUMNS.breakMinutes]) || 0,
    note: String(row[ALBYTE_ATTENDANCE_COLUMNS.note] || '').trim(),
    autoFlag: String(row[ALBYTE_ATTENDANCE_COLUMNS.autoFlag] || '').trim(),
    log: parseAlbyteAttendanceLog_(row[ALBYTE_ATTENDANCE_COLUMNS.log]),
    createdAt: parseDateValue_(row[ALBYTE_ATTENDANCE_COLUMNS.createdAt]),
    updatedAt: parseDateValue_(row[ALBYTE_ATTENDANCE_COLUMNS.updatedAt])
  };
}

function readAlbyteAttendanceRowFor_(staffId, dateKey, options){
  const sheet = options && options.sheet ? options.sheet : ensureAlbyteAttendanceSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const width = ALBYTE_ATTENDANCE_SHEET_HEADER.length;
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const targetId = String(staffId || '').trim();
  const targetDate = String(dateKey || '').trim();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (String(row[ALBYTE_ATTENDANCE_COLUMNS.staffId] || '').trim() !== targetId) continue;
    if (String(row[ALBYTE_ATTENDANCE_COLUMNS.date] || '').trim() !== targetDate) continue;
    const parsed = parseAlbyteAttendanceRow_(row, i + 2);
    if (!parsed.id) {
      parsed.id = Utilities.getUuid();
      sheet.getRange(parsed.rowIndex, ALBYTE_ATTENDANCE_COLUMN_INDEX.id).setValue(parsed.id);
    }
    return parsed;
  }
  return null;
}

function buildAlbytePortalState_(staffRecord){
  const tz = getConfig('timezone') || 'Asia/Tokyo';
  const now = new Date();
  const todayKey = fmtDate(now, tz);
  const weekday = getWeekdaySymbol_(now, tz);
  const attendance = readAlbyteAttendanceRowFor_(staffRecord.id, todayKey);
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
      record: null
    },
    presets: ALBYTE_BREAK_MINUTES_PRESETS.slice(),
    limits: {
      break: {
        max: ALBYTE_MAX_BREAK_MINUTES,
        step: ALBYTE_BREAK_STEP_MINUTES
      }
    }
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
      name: staff.name
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

      return buildAlbyteSuccessResponse_(staff, {});
    });
  });
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
  return [new Date(), String(pid), type, msg, '', metaStr];
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

function readNewsRows_(){
  const s = sh('News');
  const lr = s.getLastRow();
  if (lr < 2) return [];
  const width = Math.min(6, s.getLastColumn());
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
    const meta = parseNewsMetaValue_(metaRaw);
    const ts = Number.isFinite(tsCandidate) ? tsCandidate : 0;
    rows.push({
      ts,
      when: whenText,
      rowNumber,
      pid: normalizedPid,
      type: typeText,
      message: messageText,
      meta,
      cleared: String(disp[4] != null ? disp[4] : raw[4] || '').trim()
    });
  }
  return rows;
}

function fetchNewsRowsForPid_(normalized){
  if (!normalized) return [];
  return readNewsRows_()
    .filter(row => !row.cleared && row.pid === normalized)
    .map(row => ({
      ts: row.ts,
      when: row.when,
      type: row.type,
      message: row.message,
      meta: row.meta,
      rowNumber: row.rowNumber,
      pid: row.pid
    }));
}

function fetchGlobalNewsRows_(){
  return readNewsRows_()
    .filter(row => !row.cleared && !row.pid)
    .map(row => ({
      ts: row.ts,
      when: row.when,
      type: row.type,
      message: row.message,
      meta: row.meta,
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
  sheet.getRange(start, 1, rows.length, 6).setValues(rows);
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
        s.getRange(2+i,5).setValue('1');
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
      s.getRange(2 + i, 5).setValue('1');
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
  const width = Math.min(6, s.getMaxColumns());
  const vals = s.getRange(2, 1, lr - 1, width).getValues();
  const matchPid = String(pid || '').trim();
  const normalizedPid = normId_(matchPid);
  const filterMessage = options && options.messageContains ? String(options.messageContains) : '';
  const filterMetaType = options && options.metaType ? String(options.metaType).trim() : '';
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
      let resolvedType = '';
      if (meta && typeof meta === 'object' && meta.type != null) {
        resolvedType = String(meta.type);
      } else if (typeof meta === 'string') {
        resolvedType = meta;
      }
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
        if (String(actual) !== String(expected)) {
          metaOk = false;
        }
      });
      if (!metaOk) continue;
    }
    s.getRange(rowNumber, 5).setValue('1');
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
    s.getRange(2 + idx, clearedCol + 1).setValue('1');
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
    if (meta && typeof meta === 'object' && meta.type === 'consent_reminder') {
      existingKeys.add(row.pid + '|' + expiryKey);
      return;
    }
    if (meta && typeof meta === 'object' && meta.type === 'consent_doctor_report') {
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
            type: 'consent_doctor_report',
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
          metaType: 'consent_doctor_report',
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

    return {
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
  }, PATIENT_CACHE_TTL_SECONDS);
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
            addNews('同意','通院日未定です。後日確認してください。');
            consentReminderPushed = true;
          } else {
            const visitPlanDate = job.visitPlanDate ? String(job.visitPlanDate).trim() : '';
            const followupMessageBase = '通院日が近づいています。ご利用者様に声かけをしてください。';
            const followupMessage = visitPlanDate
              ? `${followupMessageBase}（通院予定：${visitPlanDate}）`
              : followupMessageBase;
            const meta = { type: 'consent_handout_followup' };
            if (visitPlanDate) {
              meta.visitPlanDate = visitPlanDate;
            }
            addNews('同意', followupMessage, meta);
          }
        }
      }
      if (job.consentUndecided && !consentReminderPushed){
        addNews('同意','通院日未定です。後日確認してください。');
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
    const width = Math.min(TREATMENT_SHEET_HEADER.length, s.getMaxColumns());
    const vals = s.getRange(2, 1, lr - 1, width).getValues();
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0);
    const end   = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);

    const out = [];
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      const ts = r[0];
      const id = normId_(r[1]);
      if (id !== normalized) continue;
      const d = ts instanceof Date ? ts : new Date(ts);
      if (isNaN(d.getTime())) continue;
      if (d < start || d > end) continue;
      const categoryCell = width >= 8 ? r[7] : '';
      const categoryLabel = String(categoryCell || '');
      const categoryKey = mapTreatmentCategoryCellToKey_(categoryLabel);
      out.push({
        row: 2 + i,
        when: Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm'),
        note: String(r[2] || ''),
        email: String(r[3] || ''),
        category: categoryLabel,
        categoryKey
      });
    }
    return out.reverse();
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
    invalidatePatientCaches_(pid, { header: true, treatments: true });
  }

  return { ok: true, updatedRow: row, newNote };
}

function deleteTreatmentRow(row){
  const s=sh('施術録'); const lr = s.getLastRow();
  if(row<=1 || row>lr) throw new Error('行が不正です');
  const maxCols = s.getMaxColumns();
  const width = Math.min(TREATMENT_SHEET_HEADER.length, maxCols);
  const rowVals = s.getRange(row, 1, 1, width).getValues()[0];
  const treatmentId = width >= 7 ? String(rowVals[6] || '').trim() : '';
  const pid = String(rowVals[1] || '').trim();
  s.deleteRow(row);
  if (treatmentId) clearNewsByTreatment_(treatmentId);
  log_('施術削除', '(row:'+row+')', '');
  if (pid) {
    invalidatePatientCaches_(pid, { header: true, treatments: true });
  }
  return true;
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
  const meta = options && options.meta ? options.meta : null;
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

  invalidatePatientCaches_(pid, { header: true, treatments: true });
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

function calculateBillingForMonth(ym) {
  const parsed = parseBillingMonth_(ym);
  if (!parsed) throw new Error('請求月は YYYY-MM 形式で指定してください');

  const treatment = sh('施術録');
  const patients = sh('患者情報');

  const patientLastCol = patients.getLastColumn();
  const patientLastRow = patients.getLastRow();
  const patientHead = patients.getRange(1, 1, 1, patientLastCol).getDisplayValues()[0];
  const cRec = resolveColByLabels_(patientHead, LABELS.recNo, '施術録番号');
  const cName = resolveColByLabels_(patientHead, LABELS.name, '名前');
  const cShare = resolveColByLabels_(patientHead, LABELS.share, '負担割合');

  const patientValues = patientLastRow > 1
    ? patients.getRange(2, 1, patientLastRow - 1, patientLastCol).getDisplayValues()
    : [];
  const patientMap = {};
  patientValues.forEach(row => {
    const rec = String(row[cRec - 1] || '').trim();
    if (!rec) return;
    const shareRaw = row[cShare - 1] || '';
    const shareRatio = normalizeBurdenRatio_(shareRaw);
    const shareDisp = shareRatio ? toBurdenDisp_(shareRatio) : shareRaw;
    patientMap[rec] = {
      name: row[cName - 1] || '',
      shareDisplay: shareDisp,
      shareRatio
    };
  });

  const treatmentLastRow = treatment.getLastRow();
  const counts = {};
  if (treatmentLastRow >= 2) {
    const treatmentValues = treatment.getRange(2, 1, treatmentLastRow - 1, 6).getValues();
    const start = new Date(parsed.year, parsed.month - 1, 1, 0, 0, 0);
    const end = new Date(parsed.year, parsed.month, 0, 23, 59, 59);
    treatmentValues.forEach(row => {
      const timestamp = row[0];
      const rec = String(row[1] || '').trim();
      if (!rec) return;
      const date = timestamp instanceof Date ? timestamp : new Date(timestamp);
      if (isNaN(date.getTime())) return;
      if (date >= start && date <= end) {
        counts[rec] = (counts[rec] || 0) + 1;
      }
    });
  }

  const unit = APP.BASE_FEE_YEN || 4170;
  const rows = Object.keys(counts)
    .sort((a, b) => (parseInt(a, 10) || 0) - (parseInt(b, 10) || 0))
    .map(rec => {
      const info = patientMap[rec] || { name: '', shareDisplay: '', shareRatio: null };
      const totalCount = counts[rec];
      const amount = info.shareRatio != null ? Math.round(totalCount * unit * info.shareRatio) : null;
      return {
        recNo: rec,
        name: info.name,
        totalCount,
        shareDisplay: info.shareDisplay,
        shareRatio: info.shareRatio,
        amount
      };
    });

  return {
    ym: parsed.ym,
    year: parsed.year,
    month: parsed.month,
    sheetName: '請求集計_' + parsed.sheetSuffix,
    unitFee: unit,
    rows
  };
}

function generateBillingAggregationSheet(ym) {
  const result = calculateBillingForMonth(ym);
  const wb = ss();
  let sheet = wb.getSheetByName(result.sheetName);
  if (!sheet) {
    sheet = wb.insertSheet(result.sheetName);
  } else {
    sheet.clear();
  }

  const header = ['施術録番号', '患者様氏名', '合計施術回数', '負担割合', '請求金額'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (result.rows.length) {
    const values = result.rows.map(row => [
      row.recNo,
      row.name,
      row.totalCount,
      row.shareDisplay,
      row.amount != null ? row.amount : ''
    ]);
    sheet.getRange(2, 1, values.length, header.length).setValues(values);
  }

  return { sheetName: result.sheetName, rowCount: result.rows.length };
}

function rebuildInvoiceForMonth_(year, month){
  const ym = String(year) + '-' + String(month).padStart(2, '0');
  const result = calculateBillingForMonth(ym);

  const ssb = ss();
  const outName = year + '年' + month + '月分';
  let out = ssb.getSheetByName(outName);
  if(!out) out = ssb.insertSheet(outName); else out.clear();
  out.getRange(1,1,1,4).setValues([['施術録番号','患者様氏名','合計施術回数','負担割合']]);

  if (result.rows.length) {
    const rows = result.rows.map(r => [r.recNo, r.name, r.totalCount, r.shareDisplay]);
    out.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}
function rebuildInvoiceForCurrentMonth(){
  const now=new Date(); rebuildInvoiceForMonth_(now.getFullYear(), now.getMonth()+1);
}

function promptBillingAggregation(){
  const ui = SpreadsheetApp.getUi();
  while (true) {
    const response = ui.prompt('請求集計', '請求月 (YYYY-MM) を入力してください。', ui.ButtonSet.OK_CANCEL);
    const button = response.getSelectedButton();
    if (button !== ui.Button.OK) return;

    const value = response.getResponseText();
    const parsed = parseBillingMonth_(value);
    if (!parsed) {
      ui.alert('請求月は YYYY-MM 形式で入力してください。');
      continue;
    }

    try {
      const result = generateBillingAggregationSheet(parsed.ym);
      ui.alert('請求集計', '請求月 ' + parsed.ym + ' の集計が完了しました。\n出力先シート: ' + result.sheetName, ui.ButtonSet.OK);
    } catch (e) {
      Logger.log('[promptBillingAggregation] ' + e);
      ui.alert('請求集計に失敗しました: ' + (e && e.message ? e.message : e));
    }
    return;
  }
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
  const metaMatches = { type: 'consent_doctor_report' };
  if (payload && payload.consentExpiry) {
    metaMatches.consentExpiry = String(payload.consentExpiry);
  }
  const options = {
    metaType: 'consent_doctor_report',
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
  const view = e.parameter ? (e.parameter.view || 'welcome') : 'welcome';
  let templateFile = '';

  switch(view){
    case 'intake':       templateFile = 'intake'; break;
    case 'visit':        templateFile = 'intake'; break;
    case 'intake_list':  templateFile = 'intake_list'; break;
    case 'admin':        templateFile = 'admin'; break;
    case 'attendance':   templateFile = 'attendance'; break;
    case 'vacancy':      templateFile = 'vacancy'; break;
    case 'albyte':       templateFile = 'albyte'; break;
    case 'record':       templateFile = 'app'; break;   // ★ app.html を record として表示
    case 'report':       templateFile = 'report'; break;
    default:             templateFile = 'welcome'; break;
  }

  const t = HtmlService.createTemplateFromFile(templateFile);

  // ここでURLを渡す
  t.baseUrl = ScriptApp.getService().getUrl();

  // 患者ID（?id=XXXX）をテンプレートに渡す
  if (e.parameter && e.parameter.id) {
    t.patientId = e.parameter.id;
  } else {
    t.patientId = "";
  }

  if(e.parameter && e.parameter.lead) t.lead = e.parameter.lead;

  return t.evaluate()
           .setTitle('受付アプリ')
           .addMetaTag('viewport','width=device-width, initial-scale=1.0');
}

/***** メニュー *****/
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('請求')
    .addItem('今月の集計（回数+負担割合）','rebuildInvoiceForCurrentMonth')
    .addToUi();
  ui.createMenu('請求集計')
    .addItem('請求月を指定して集計','promptBillingAggregation')
    .addToUi();
  ui.createMenu('勤怠管理')
    .addItem('勤怠データを今すぐ同期','runVisitAttendanceSyncJobFromMenu')
    .addItem('日次同期トリガーを確認','ensureVisitAttendanceSyncTriggerFromMenu')
    .addToUi();
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
  const values = sheet.getRange(2,1,lastRow-1,6).getValues();
  values.forEach(row => {
    const ts = row[0];
    const email = normalizeEmailKey_(row[3]);
    if (!email) return;
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
  const sheet = sh('施術録');
  const lastRow = sheet.getLastRow();
  const summary = {
    date: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
    staffProcessed: 0,
    posted: 0,
    skipped: 0,
    totalTreatments: 0
  };
  if (lastRow < 2) {
    Logger.log('[sendDailySummaryToChat] 施術録にデータがありません');
    return summary;
  }

  const values = sheet.getRange(2,1,lastRow-1,6).getValues();
  const byStaff = new Map();
  const patientIds = new Set();

  values.forEach(row => {
    const ts = row[0];
    const rawId = row[1];
    const emailRaw = String(row[3] || '').trim();
    if (!emailRaw) return;
    const when = ts instanceof Date ? ts : new Date(ts);
    if (isNaN(when.getTime()) || when < start || when >= end) return;

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
      UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ text: message })
      });
      summary.posted += 1;
    } catch (err) {
      Logger.log(`[sendDailySummaryToChat] 送信失敗 staff=${entry.email} err=${err}`);
      summary.skipped += 1;
    }
  });

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
  try{
    const s = sh('通知設定'); const lr=s.getLastRow(); if(lr<2) return false;
    const vals = s.getRange(2,1,lr-1,3).getDisplayValues(); // [スタッフメール,WebhookURL,管理者]
    const me = (Session.getActiveUser()||{}).getEmail() || '';
    return vals.some(r => (String(r[0]||'').toLowerCase()===me.toLowerCase()) && String(r[2]||'').toUpperCase()==='TRUE');
  }catch(e){ return false; }
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

  invalidatePatientCaches_(pid, { header: true, treatments: true });
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
    if (!categoryKey) {
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

    const presetLabel = String(payload?.presetLabel || '').trim();
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
      invalidatePatientCaches_(pid, { header: true, treatments: true });
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
      if (group === 'mixed') pendingEntry.counts.mixed += 1;
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
    map.set(email, {
      email,
      displayName: String(row[1] || '').trim(),
      quotaDays: quota
    });
  });
  return map;
}

function resolveAnnualPaidLeaveQuota_(email){
  const normalized = normalizeEmailKey_(email);
  if (!normalized) return DEFAULT_ANNUAL_PAID_LEAVE_DAYS;
  const settings = readVisitAttendanceStaffSettings_();
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

function readVisitAttendanceRecordsForEmail_(email, options){
  const normalizedEmail = normalizeEmailKey_(email);
  if (!normalizedEmail) return [];
  const opts = options || {};
  const tz = opts.tz || Session.getScriptTimeZone() || 'Asia/Tokyo';
  const startDate = opts.startDate instanceof Date ? new Date(opts.startDate.getFullYear(), opts.startDate.getMonth(), opts.startDate.getDate()) : null;
  const endDate = opts.endDate instanceof Date ? new Date(opts.endDate.getFullYear(), opts.endDate.getMonth(), opts.endDate.getDate()) : null;
  const startMs = startDate ? startDate.getTime() : null;
  const endMs = endDate ? endDate.getTime() : null;
  const sheet = ensureVisitAttendanceSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const width = Math.min(VISIT_ATTENDANCE_SHEET_HEADER.length, sheet.getMaxColumns());
  const range = sheet.getRange(2, 1, lastRow - 1, width);
  const values = range.getValues();
  const displays = range.getDisplayValues();
  const weekdays = ['日','月','火','水','木','金','土'];
  const results = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const display = displays[i];
    const rowEmail = normalizeEmailKey_(row[1] || display[1]);
    if (!rowEmail || rowEmail !== normalizedEmail) continue;
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

    results.push({
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
  results.sort((a, b) => a.date.localeCompare(b.date));
  return results;
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
  const quotaDays = resolveAnnualPaidLeaveQuota_(normalizedEmail);
  const usage = calculatePaidLeaveUsageForYear_(normalizedEmail, currentYear, tz);
  const remainingDays = Math.max(0, quotaDays - (usage.usedDays || 0));
  const paidLeaveSummary = {
    year: currentYear,
    quotaDays,
    usedDays: usage.usedDays || 0,
    remainingDays,
    requiredDays: 5
  };

  return {
    ok: true,
    user: {
      email: normalizedEmail,
      displayName: resolveStaffDisplayName_(normalizedEmail),
      isAdmin
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
    paidLeave: paidLeaveSummary
  };
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
  let targetDate = data.date || data.targetDate;
  if (targetDate instanceof Date && !isNaN(targetDate.getTime())) {
    targetDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
  } else {
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
    targetDate = new Date(year, monthIndex, day);
  }
  if (!(targetDate instanceof Date) || isNaN(targetDate.getTime())) {
    throw new Error('日付の解析に失敗しました');
  }

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
      let workMinutes = Number(data.workMinutes);
      if (!Number.isFinite(workMinutes) || workMinutes <= 0) {
        workMinutes = PAID_LEAVE_DEFAULT_WORK_MINUTES;
      }
      createVisitAttendanceRecord({
        email: targetEmail,
        date: targetDay,
        startMinutes: VISIT_ATTENDANCE_WORK_START_MINUTES,
        workMinutes,
        restMinutes: 0,
        leaveType: VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE,
        source: VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE,
        flag: VISIT_ATTENDANCE_REQUEST_TYPE_PAID_LEAVE,
        breakdown: '有給',
        isHourlyStaff: false,
        isDailyStaff: false
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
