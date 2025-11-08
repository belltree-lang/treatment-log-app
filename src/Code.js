/***** ── 設定 ─────────────────────────────────*****/
const APP = {
  // Driveに保存するPDFの親フォルダID（空でも可：スプレッドシートと同じ階層に保存）
  PARENT_FOLDER_ID: '1VAv9ZOLB7A__m8ErFDPhFHvhpO21OFPP',
  // 正本スプレッドシート（患者情報のブック）。空なら「現在のスプレッドシート」を使う
  SSID: '1ajnW9Fuvu0YzUUkfTmw0CrbhrM3lM5tt5OA1dK2_CoQ',
  BASE_FEE_YEN: 4170,
  // 社内ドメイン制限（空＝無効）
  ALLOWED_DOMAIN: '',   // 例 'belltree1102.com'

  // OpenAI（任意・未設定ならローカル整形へフォールバック）
  OPENAI_ENDPOINT: 'https://api.openai.com/v1/chat/completions',
  OPENAI_MODEL: 'gpt-4o-mini',
};

const CLINICAL_METRICS = [
  { id: 'pain_vas',      label: '痛みVAS',           unit: '/10', min: 0,   max: 10,  step: 0.5, description: '主観的疼痛スケール（0=痛みなし, 10=最大）' },
  { id: 'rom_knee_flex', label: '膝屈曲ROM',         unit: '°',   min: 0,   max: 150, step: 1,   description: '膝関節屈曲の可動域' },
  { id: 'rom_knee_ext',  label: '膝伸展ROM',         unit: '°',   min: -20, max: 10,  step: 1,   description: '膝関節伸展の可動域（マイナスは屈曲拘縮）' },
  { id: 'walk_distance', label: '歩行距離（6MWT）', unit: 'm',   min: 0,   max: 600, step: 5,   description: '6分間歩行距離などの歩行パフォーマンス' },
];

const AUX_SHEETS_INIT_KEY = 'AUX_SHEETS_INIT_V202501';
const PATIENT_CACHE_TTL_SECONDS = 90;
const PATIENT_CACHE_KEYS = {
  header: pid => 'patient:header:' + normId_(pid),
  news: pid => 'patient:news:' + normId_(pid),
  treatments: pid => 'patient:treatments:' + normId_(pid),
};
const GLOBAL_NEWS_CACHE_KEY = 'patient:news:__global__';

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
    const need = ['施術録','患者情報','News','フラグ','予定','操作ログ','定型文','添付索引','年次確認','ダッシュボード','臨床指標','AI報告書'];
    need.forEach(n => { if (!wb.getSheetByName(n)) wb.insertSheet(n); });

    const ensureHeader = (name, header) => {
      const s = wb.getSheetByName(name);
      if (s.getLastRow() === 0) s.appendRow(header);
    };

    // 既存タブ
    ensureHeader('施術録',   ['タイムスタンプ','施術録番号','所見','メール','最終確認','名前','treatmentId']);
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

    upgradeHeader('施術録', ['タイムスタンプ','施術録番号','所見','メール','最終確認','名前','treatmentId']);
    upgradeHeader('News',   ['TS','患者ID','種別','メッセージ','cleared','meta']);
    ensureHeader('フラグ',   ['患者ID','status','pauseUntil']);
    ensureHeader('予定',     ['患者ID','種別','予定日','登録者']);
    ensureHeader('操作ログ', ['TS','操作','患者ID','詳細','実行者']);
    ensureHeader('定型文',   ['カテゴリ','ラベル','文章']);
    ensureHeader('添付索引', ['TS','患者ID','月','ファイル名','FileId','種別','登録者']);
    ensureHeader('AI報告書', ['TS','患者ID','範囲','対象','status','special']);

    // 年次確認タブ（未作成時はヘッダだけ用意）
    ensureHeader('年次確認', ['患者ID','年','確認日','担当者メール']);

    // ダッシュボード（Index）タブ
    ensureHeader('ダッシュボード', [
      '患者ID','氏名','同意年月日','次回期限','期限ステータス',
      '担当者(60d)','最終施術日','年次要確認','休止','ミュート解除予定','負担割合整合'
    ]);

    ensureHeader('臨床指標', ['TS','患者ID','指標ID','値','メモ','登録者']);

    props.setProperty(AUX_SHEETS_INIT_KEY, '1');
  } finally {
    if (locked) {
      lock.releaseLock();
    }
  }
}

function getClinicalMetricDefinitions(){
  return CLINICAL_METRICS.map(m => ({
    id: m.id,
    label: m.label,
    unit: m.unit || '',
    min: m.min,
    max: m.max,
    step: m.step || 1,
    description: m.description || ''
  }));
}

function getClinicalMetricDef_(id){
  return CLINICAL_METRICS.find(m => m.id === id) || null;
}

function ensureClinicalMetricSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName('臨床指標');
  if (!sheet) {
    const conflict = wb.getSheets().find(s => /^臨床指標[_\-]?conflict/i.test(s.getName()));
    if (conflict) {
      conflict.setName('臨床指標');
      sheet = conflict;
    } else {
      sheet = wb.insertSheet('臨床指標');
    }
  }

  wb.getSheets()
    .filter(s => s !== sheet && /^臨床指標[_\-]?conflict/i.test(s.getName()))
    .forEach(s => {
      if (s.getLastRow() <= 1) {
        wb.deleteSheet(s);
      } else {
        Logger.log(`[ensureClinicalMetricSheet_] 競合シートを検出: ${s.getName()}`);
      }
    });

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['TS','患者ID','指標ID','値','メモ','登録者']);
  }
  return sheet;
}

function ensureAiReportSheet_(){
  ensureAuxSheets_();
  const wb = ss();
  let sheet = wb.getSheetByName('AI報告書');
  if (!sheet) {
    sheet = wb.insertSheet('AI報告書');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['TS','患者ID','範囲','対象','status','special']);
  }
  return sheet;
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
    if (filterMetaType) {
      const metaRaw = width >= 6 ? vals[i][5] : '';
      const meta = parseNewsMetaValue_(metaRaw);
      let resolvedType = '';
      if (meta && typeof meta === 'object' && meta.type != null) {
        resolvedType = String(meta.type);
      } else if (typeof meta === 'string') {
        resolvedType = meta;
      }
      if (resolvedType !== filterMetaType) continue;
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
  existing.forEach(row => {
    if (row.cleared) return;
    if (!row.pid) return;
    if (String(row.type || '').trim() !== '同意') return;
    const message = String(row.message || '').trim();
    if (message !== '同意書受渡が必要です') return;
    const meta = row.meta;
    let expiryKey = '';
    if (meta && typeof meta === 'object' && meta.consentExpiry) {
      expiryKey = String(meta.consentExpiry);
    }
    existingKeys.add(row.pid + '|' + expiryKey);
  });

  const toInsert = [];
  const insertedKeys = new Set();
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
    const reminderDate = new Date(expiryDate.getTime() - 30 * dayMs);
    reminderDate.setHours(0, 0, 0, 0);
    const diffDays = Math.round((reminderDate.getTime() - todayStart.getTime()) / dayMs);
    if (diffDays !== 0) continue;
    const key = pidNormalized + '|' + expiryStr;
    if (existingKeys.has(key) || insertedKeys.has(key)) {
      continue;
    }
    const pidForNews = String(pidRaw || '').trim();
    if (!pidForNews) continue;
    const meta = {
      source: 'auto',
      type: 'consent_reminder',
      consentExpiry: expiryStr,
      reminderDate: Utilities.formatDate(reminderDate, tz, 'yyyy-MM-dd')
    };
    toInsert.push(formatNewsRow_(pidForNews, '同意', '同意書受渡が必要です', meta));
    insertedKeys.add(key);
  }

  if (toInsert.length) {
    pushNewsRows_(toInsert);
  }
  return { ok: true, scanned: rows.length, inserted: toInsert.length };
}

function checkConsentExpiration(){
  return checkConsentExpiration_();
}

/***** ステータス（休止/中止） *****/
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
  return vals.map(r=>({cat:r[0],label:r[1],text:r[2]}));
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

function normalizeClinicalMetricTimestamp_(value){
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  if (value) {
    const parsed = parseDateTimeFlexible_(value, tz) || parseDateFlexible_(value);
    if (parsed && !isNaN(parsed.getTime())) return parsed;
  }
  return new Date();
}

function recordClinicalMetrics_(patientId, metrics, whenStr, user){
  const pid = String(patientId || '').trim();
  if (!pid) {
    Logger.log('臨床指標は未入力です（保存はスキップしました） [pid=]');
    return;
  }

  if (!Array.isArray(metrics) || !metrics.length) {
    Logger.log(`臨床指標は未入力です（保存はスキップしました） [pid=${pid}]`);
    return;
  }

  const sheet = ensureClinicalMetricSheet_();
  const rows = [];
  const timestamp = normalizeClinicalMetricTimestamp_(whenStr);
  const owner = user ? String(user).trim() : '';

  metrics.forEach(item => {
    if (!item) return;
    const metricId = String(item.metricId || item.id || '').trim();
    const def = getClinicalMetricDef_(metricId);
    if (!def) return;
    const rawVal = item.value != null ? Number(item.value) : NaN;
    if (!isFinite(rawVal)) return;
    const note = item.note ? String(item.note).trim() : '';
    rows.push([
      timestamp,
      pid,
      metricId,
      rawVal,
      note,
      owner
    ]);
  });

  if (!rows.length) {
    Logger.log(`臨床指標は未入力です（保存はスキップしました） [pid=${pid}]`);
    return;
  }

  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, 6).setValues(rows);
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
      const addNews = (type, message) => {
        newsRows.push(formatNewsRow_(pid, type, message, treatmentMeta));
      };

      // News / 同意日 / 負担割合 / 予定登録など重い処理をここでまとめて実行
      let consentReminderPushed = false;
      if (job.presetLabel){
        if (job.presetLabel.indexOf('再同意取得確認') >= 0){
          const today = Utilities.formatDate(new Date(), tz,'yyyy-MM-dd');
          updateConsentDate(pid, today, treatmentMeta ? { meta: treatmentMeta } : undefined);
        }
        if (job.presetLabel.indexOf('同意書受渡') >= 0){
          addNews('再同意','同意書を受け渡し');
          if (job.consentUndecided){
            addNews('同意','同意日未定です。後日確認してください。');
            consentReminderPushed = true;
          }
        }
      }
      if (job.consentUndecided && !consentReminderPushed){
        addNews('同意','同意日未定です。後日確認してください。');
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
    const width = Math.min(7, s.getMaxColumns());
    const vals = s.getRange(2, 1, lr - 1, width).getValues(); // A..G
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
      out.push({
        row: 2 + i,
        when: Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm'),
        note: String(r[2] || ''),
        email: String(r[3] || '')
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
  const width = Math.min(7, maxCols);
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

function parseClinicalMetricTimestamp_(value){
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const parsed = parseDateFlexible_(value);
  return parsed || new Date(value);
}

function listClinicalMetricSeries(pid, startDate, endDate){
  const sheet = ensureClinicalMetricSheet_();
  const lr = sheet.getLastRow();
  if (lr < 2) return { metrics: [] };

  const vals = sheet.getRange(2, 1, lr - 1, 6).getValues();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const normalizedPid = String(normId_(pid));
  const startTs = startDate instanceof Date ? startDate.getTime() : null;
  const endTs = endDate instanceof Date ? endDate.getTime() : null;
  const map = {};

  vals.forEach(row => {
    const ts = parseClinicalMetricTimestamp_(row[0]);
    const rowPid = String(normId_(row[1]));
    if (!normalizedPid || rowPid !== normalizedPid) return;
    if (!(ts instanceof Date) || isNaN(ts.getTime())) return;
    const ms = ts.getTime();
    if (startTs != null && ms < startTs) return;
    if (endTs != null && ms > endTs) return;
    const metricId = String(row[2] || '').trim();
    const def = getClinicalMetricDef_(metricId);
    if (!def) return;
    const value = Number(row[3]);
    if (!isFinite(value)) return;
    const note = row[4] || '';
    const user = row[5] || '';
    const dispDate = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
    if (!map[metricId]) map[metricId] = [];
    map[metricId].push({ date: dispDate, value, note: note ? String(note) : '', user: String(user || '') });
  });

  const defs = getClinicalMetricDefinitions();
  const metrics = defs
    .map(def => ({
      id: def.id,
      label: def.label,
      unit: def.unit || '',
      min: def.min,
      max: def.max,
      step: def.step,
      description: def.description || '',
      points: (map[def.id] || []).sort((a, b) => (a.date > b.date ? 1 : a.date < b.date ? -1 : 0))
    }))
    .filter(m => m.points.length);

  return { metrics };
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
  const end = new Date();
  end.setHours(23, 59, 59, 999);
  let start = null;
  let label = '全期間';

  switch (rangeKey) {
    case '1m':
    case '2m':
    case '3m': {
      const months = Number(rangeKey.replace('m', ''));
      label = `直近${months}か月`;
      start = new Date(end.getTime());
      start.setHours(0, 0, 0, 0);
      start.setMonth(start.getMonth() - months);
      break;
    }
    default:
      label = '全期間';
      start = null;
      break;
  }

  if (start) {
    start = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 0, 0, 0, 0);
  }

  return { startDate: start, endDate: end, label };
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
    const wb = ss();
    const sheet = wb.getSheetByName('同意書');
    if (!sheet) return '';
    const lr = sheet.getLastRow();
    if (lr < 2) return '';
    const lc = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lc).getDisplayValues()[0].map(h => String(h || '').trim());

    const findCol = (predicate) => {
      for (let i = 0; i < headers.length; i++) {
        if (predicate(headers[i])) return i + 1;
      }
      return null;
    };

    const pidCol = findCol(h => /患者|施術録|ID|番号/.test(h));
    const contentCol = findCol(h => h.indexOf('同意') >= 0 && (h.indexOf('内容') >= 0 || h.indexOf('事項') >= 0 || h.indexOf('概要') >= 0 || h.indexOf('文') >= 0));
    if (!pidCol || !contentCol) return '';

    const vals = sheet.getRange(2, 1, lr - 1, lc).getDisplayValues();
    const target = String(normId_(pid));
    for (let i = 0; i < vals.length; i++) {
      const row = vals[i];
      if (String(normId_(row[pidCol - 1])) === target) {
        return String(row[contentCol - 1] || '').trim();
      }
    }
    return '';
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
  const metricsDigest = normalizeDoctorReportText_(context?.metricsDigest);
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
  if (metricsDigest) {
    safetySource = safetySource
      ? `${safetySource} / 臨床指標：${metricsDigest}`
      : `臨床指標：${metricsDigest}`;
  }
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
  // AI抽出部分：リスク・体調管理＋末尾に必ず「同意内容に沿った施術を継続しております。」
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

function buildMetricDigestForSummary_(metrics){
  if (!Array.isArray(metrics) || !metrics.length) return '';
  const lines = [];
  metrics.forEach(metric => {
    const pts = Array.isArray(metric?.points) ? metric.points : [];
    if (!pts.length) return;
    const last = pts[pts.length - 1];
    const val = last ? `${last.date || ''} ${last.value}${metric.unit || ''}` : '';
    if (val) lines.push(`${metric.label}: 最新 ${val}`);
  });
  return lines.join(' / ');
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
    'all': 'all'
  };
  if (map[raw]) return map[raw];
  if (map[lower]) return map[lower];
  const match = raw.match(/直近\s*(\d+)\s*か?月/);
  if (match) {
    const months = Math.max(1, Number(match[1] || 1));
    return `${months}m`;
  }
  return raw;
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

  const metrics = context?.metrics && Array.isArray(context.metrics.metrics)
    ? context.metrics.metrics
    : [];
  if (metrics.length) {
    lines.push('【臨床指標（最新値）】');
    metrics.forEach(metric => {
      const points = Array.isArray(metric?.points) ? metric.points : [];
      if (!points.length) return;
      const last = points[points.length - 1];
      const value = last?.value != null ? `${last.value}${metric?.unit || ''}` : '';
      const date = String(last?.date || '').trim();
      const note = String(last?.note || '').trim();
      const desc = [date, value, note].filter(Boolean).join(' / ');
      const label = metric?.label || metric?.id;
      lines.push(`- ${label}: ${desc}`);
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
  const range = resolveIcfSummaryRange_(rangeKey || 'all');
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

  const header = {
    hospital: source.hospital,
    doctor: source.doctor,
    name: source.name,
    birth: source.birth,
    patientId: patientId
  };

  const context = {
    consentText: source.consent,
    frequencyLabel: source.frequencyLabel,
    rangeLabel: range.label,
    metricsDigest: source.metricsDigest,
    notes: source.notes,
    handovers: source.handovers,
    metrics: source.metrics
  };

  // ★ AIに直接投げる
  const aiResult = composeAiReportViaOpenAI_(header, context, audienceMeta.key);
  const text = (aiResult && typeof aiResult === 'object') ? aiResult.text : String(aiResult || '');

  return {
    ok: true,
    usedAi: true,
    audience: audienceMeta.key,
    audienceLabel: audienceMeta.label,
    text: text,
    meta: {
      patientFound: true,
      rangeLabel: range.label,
      noteCount: Array.isArray(source.notes) ? source.notes.length : 0,
      handoverCount: Array.isArray(source.handovers) ? source.handovers.length : 0,
      metricCount: source.metrics?.metrics?.length || 0
    }
  };
}
/***** OpenAI で AI レポート生成 *****/
function composeAiReportViaOpenAI_(header, context, audienceKey) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY が設定されていません。');
  }

  const prompt = buildReportPrompt_(header, context, audienceKey);

  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: 'gpt-4o-mini', // または gpt-4o / gpt-4.1 など
    messages: [
      { role: 'system', content: 'あなたは鍼灸マッサージ院の施術経過を医師・ケアマネ・家族向けに報告する専門アシスタントです。' },
      { role: 'user', content: prompt }
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
  const roleLabel = audienceKey === 'doctor'
    ? '医師'
    : audienceKey === 'caremanager'
      ? 'ケアマネジャー'
      : 'ご家族';

  return `
【病院名】${header.hospital || '—'}
【担当医名】${header.doctor || '—'}
【患者氏名】${header.name || '—'}
【生年月日】${header.birth || '—'}
【同意内容】${context.consentText || '—'}
【施術頻度】${context.frequencyLabel || '—'}

${roleLabel}向けに患者様の状態・経過をまとめてください。
必ず「同意内容に沿った施術を継続しております。」という一文を含めてください。

参考情報：
- Notes: ${JSON.stringify(context.notes || [])}
- Handovers: ${JSON.stringify(context.handovers || [])}
- Metrics: ${JSON.stringify(context.metrics || [])}
- 期間: ${context.rangeLabel}
`;
}


function composeAiReportLocal_(header, context, reportType){
  const audienceMeta = resolveAudienceMeta_(reportType);
  const range = context?.range || { startDate: null, endDate: new Date(), label: '全期間' };
  const sections = Array.isArray(context?.sections) ? context.sections : [];
  const source = context?.source || { header, notes: [], handovers: [], metrics: context?.metrics };
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

function recordAiReportOutcome_(header, range, audienceMeta, outcome, final){
  if (!header || !header.patientId) return;
  const sheet = ensureAiReportSheet_();
  const statusParts = [];
  statusParts.push(outcome.usedAi ? 'via=ai' : 'via=local');
  const meta = outcome.meta || {};
  if (meta.noteCount != null) statusParts.push(`notes=${meta.noteCount}`);
  if (meta.handoverCount != null) statusParts.push(`handovers=${meta.handoverCount}`);
  if (meta.metricCount != null) statusParts.push(`metrics=${meta.metricCount}`);
  if (final && final.httpCode) statusParts.push(`http=${final.httpCode}`);
  if (final && typeof final.responseLength === 'number') statusParts.push(`len=${final.responseLength}`);
  const status = statusParts.join(' | ');
  const specialText = Array.isArray(outcome.special)
    ? outcome.special.join('\n')
    : String(outcome.special || '');
  sheet.appendRow([
    new Date(),
    String(header.patientId || ''),
    range && range.label ? range.label : '',
    audienceMeta && audienceMeta.label ? audienceMeta.label : audienceMeta.key || '',
    status,
    specialText
  ]);
}

function buildIcfSource_(pid, range){
  const header = getPatientHeader(pid);
  if (!header) {
    return { patientFound: false };
  }
  const notes = getTreatmentNotesInRange_(pid, range.startDate, range.endDate);
  const handovers = getHandoversInRange_(pid, range.startDate, range.endDate);
  const metrics = listClinicalMetricSeries(pid, range.startDate, range.endDate);
  return {
    patientFound: true,
    header,
    notes,
    handovers,
    metrics
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
  const metrics = source.metrics && Array.isArray(source.metrics.metrics)
    ? source.metrics.metrics
    : [];
  const handovers = Array.isArray(source.handovers) ? source.handovers : [];
  const sectionSummary = summarizeSectionsForAudience_(audienceKey, sections);
  const metricsDigest = buildMetricDigestForSummary_(metrics);
  const handoverDigest = buildHandoverDigestForSummary_(handovers, audienceKey);

  if (audienceKey === 'doctor') {
    const context = {
      consentText: getConsentContentForPatient_(header.patientId),
      frequencyLabel: determineTreatmentFrequencyLabel_(countTreatmentsInRecentMonth_(header.patientId, range.endDate)),
      rangeLabel,
      metricsDigest
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
    if (metricsDigest) lines.push(`【臨床指標】${metricsDigest}`);
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
  if (metricsDigest) lines.push(`最新の指標：${metricsDigest}`);
  lines.push('ご不明な点があればいつでもご連絡ください。');
  return lines.join('\n');
}

/**
 * 単一オーディエンス用：医師／ケアマネ／家族 向けサマリを生成
 */
function generateAiSummaryServer(patientId, rangeKey, audience) {
  const range = resolveIcfSummaryRange_(rangeKey || 'all');
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


const patientInfo = getPatientHeader(patientId);  // ← patientId を使う

const header = {
  hospital: patientInfo?.hospital || '—',
  doctor:   patientInfo?.doctor   || '—',
  name:     patientInfo?.name     || '—',
  birth:    patientInfo?.birth    || '—',
  consent:  patientInfo?.consent  || '—',
  patientId: patientId
};

  // コンテキスト情報
  const context = {
    consentText: source.consent,
    frequencyLabel: source.frequencyLabel,
    rangeLabel: range.label,
    metricsDigest: source.metricsDigest,
    notes: source.notes,
    handovers: source.handovers,
    metrics: source.metrics
  };

  // ★ AIに直接投げる
  const aiRes = composeAiReportViaOpenAI_(header, context, audienceMeta.key);
  const text = aiRes && aiRes.text ? aiRes.text : '';

  return {
    ok: true,
    usedAi: true,
    audience: audienceMeta.key,
    audienceLabel: audienceMeta.label,
    text: text,
    meta: {
      patientFound: true,
      rangeLabel: range.label,
      noteCount: Array.isArray(source.notes) ? source.notes.length : 0,
      handoverCount: Array.isArray(source.handovers) ? source.handovers.length : 0,
      metricCount: source.metrics?.metrics?.length || 0
    }
  };
}

/**
 * 3種類まとめて生成（doctor / caremanager / family）
 */
function generateAllAiSummariesServer(patientId, rangeKey) {
  const range = resolveIcfSummaryRange_(rangeKey || 'all');
  const source = buildIcfSource_(patientId, range);

  if (!source.patientFound) {
    return {
      ok: false,
      usedAi: true,
      reports: null,
      meta: { patientFound: false, rangeLabel: range.label }
    };
  }

  // ヘッダ情報
  const header = {
    hospital: source.hospital,
    doctor: source.doctor,
    name: source.name,
    birth: source.birth,
    patientId: patientId
  };

  // コンテキスト情報
  const context = {
    consentText: source.consent,
    frequencyLabel: source.frequencyLabel,
    rangeLabel: range.label,
    metricsDigest: source.metricsDigest,
    notes: source.notes,
    handovers: source.handovers,
    metrics: source.metrics
  };

  // 3種類まとめて生成
  const doctor = composeAiReportViaOpenAI_(header, context, 'doctor')?.text || '';
  const caremanager = composeAiReportViaOpenAI_(header, context, 'caremanager')?.text || '';
  const family = composeAiReportViaOpenAI_(header, context, 'family')?.text || '';


return {
  ok: true,
  usedAi: true,
  reports: {
    doctor: {
      ok: true,
      usedAi: true,
      audience: 'doctor',
      audienceLabel: getIcfAudienceLabel_('doctor'),
      text: (doctor && typeof doctor === 'object') ? doctor.text : String(doctor || '')
    },
    caremanager: {
      ok: true,
      usedAi: true,
      audience: 'caremanager',
      audienceLabel: getIcfAudienceLabel_('caremanager'),
      text: (caremanager && typeof caremanager === 'object') ? caremanager.text : String(caremanager || '')
    },
    family: {
      ok: true,
      usedAi: true,
      audience: 'family',
      audienceLabel: getIcfAudienceLabel_('family'),
      text: (family && typeof family === 'object') ? family.text : String(family || '')
    }
  },
  meta: {
    patientFound: true,
    rangeLabel: range.label,
    noteCount: Array.isArray(source.notes) ? source.notes.length : 0,
    handoverCount: Array.isArray(source.handovers) ? source.handovers.length : 0,
    metricCount: source.metrics?.metrics?.length || 0
  }
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
    rangeLabel: reports.rangeLabel,
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

  const rangeKeyRaw = payload.range || payload.rangeKey || 'all';
  const rangeKey = normalizeAudienceRange_(rangeKeyRaw);
  return generateAiSummaryServer(patientId, rangeKey, meta.key);
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
    case 'vacancy':      templateFile = 'vacancy'; break;
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
    pid = String(payload?.patientId || '').trim();
    if (!pid) throw new Error('patientIdが空です');

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
    const row = [now, pid, merged, user, '', '', treatmentId];
    s.appendRow(row);
    markTiming('appendRow');

    if (Array.isArray(payload?.clinicalMetrics) && payload.clinicalMetrics.length) {
      recordClinicalMetrics_(pid, payload.clinicalMetrics, now, user);
      markTiming('metrics');
    }

    const job = { patientId: pid, treatmentId, treatmentTimestamp: now };
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

    if (hasFollowUp) {
      queueAfterTreatmentJob(job);
      markTiming('queueJob');
    }

    markTiming('done');
    logSubmitTreatmentTimings_(pid, treatmentId, 'ok', timings);
    timingLogged = true;

    invalidatePatientCaches_(pid, { header: true, treatments: true });
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

  if (Object.prototype.hasOwnProperty.call(payload || {}, 'clinicalMetrics')) {
    try {
      recordClinicalMetrics_(pid, payload.clinicalMetrics, now, user);
    } catch (err) {
      Logger.log('[saveHandover] 臨床指標保存エラー: ' + err);
    }
  }

  return { ok:true, fileIds };
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
