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

/***** 先頭行（見出し）の揺れに耐えるためのラベル候補群 *****/
const LABELS = {
  recNo:     ['施術録番号','施術録No','施術録NO','記録番号','カルテ番号','患者ID','患者番号'],
  name:      ['名前','氏名','患者名','お名前'],
  hospital:  ['病院名','医療機関','病院'],
  doctor:    ['医師','主治医','担当医'],
  furigana:  ['ﾌﾘｶﾞﾅ','ふりがな','フリガナ'],
  birth:     ['生年月日','誕生日','生年','生年月'],
  consent:   ['同意年月日','同意日','同意開始日','同意開始'],
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
/***** 補助タブの用意（不足時に自動生成＋ヘッダ挿入） *****/
function ensureAuxSheets_() {
  const wb = ss();
  const need = ['施術録','患者情報','News','フラグ','予定','操作ログ','定型文','添付索引','年次確認','ダッシュボード','臨床指標','AI報告書'];
  need.forEach(n => { if (!wb.getSheetByName(n)) wb.insertSheet(n); });

  const ensureHeader = (name, header) => {
    const s = wb.getSheetByName(name);
    if (s.getLastRow() === 0) s.appendRow(header);
  };

  // 既存タブ
  ensureHeader('施術録',   ['タイムスタンプ','施術録番号','所見','メール','最終確認','名前']);
  ensureHeader('News',     ['TS','患者ID','種別','メッセージ','cleared']);
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
function pushNews_(pid,type,msg){
  sh('News').appendRow([new Date(), String(pid), type, msg, '']);
}
function getNews(pid){
  const s=sh('News'); const lr=s.getLastRow(); if(lr<2) return [];
  const vals=s.getRange(2,1,lr-1,5).getDisplayValues();
  return vals.filter(r=> String(r[1])===String(pid) && !String(r[4])).map(r=>({ when:r[0], type:r[2], message:r[3] }));
}
function clearConsentRelatedNews_(pid){
  const s=sh('News'); const lr=s.getLastRow(); if(lr<2) return;
  const vals=s.getRange(2,1,lr-1,5).getValues(); // [TS,pid,type,msg,cleared]
  for (let i=0;i<vals.length;i++){
    if(String(vals[i][1])===String(pid)){
      const typ=String(vals[i][2]||'');
      if(typ.indexOf('同意')>=0 || typ.indexOf('期限')>=0 || typ.indexOf('予定')>=0){
        s.getRange(2+i,5).setValue('1');
      }
    }
  }
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
}
function markStop(pid){
  ensureAuxSheets_();
  sh('フラグ').appendRow([String(pid),'stopped','']);
  pushNews_(pid,'状態','中止に設定（以降のリマインド停止）');
  log_('中止', pid, '');
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
    consentExpiry: expiry,
    burden: shareDisp || '',
    monthly, recent,
    status: stat.status,
    pauseUntil: stat.pauseUntil
  };
}

/***** ID候補 *****/
function listPatientIds(){
  const s=sh('患者情報'); const lr=s.getLastRow(); if(lr<2) return [];
  const lc=s.getLastColumn(); const head=s.getRange(1,1,1,lc).getDisplayValues()[0];
  const cRec = getColFlexible_(head, LABELS.recNo, PATIENT_COLS_FIXED.recNo, '施術録番号');
  const vals=s.getRange(2,1,lr-1,lc).getValues();
  return vals.map(r=> normId_(r[cRec-1])).filter(Boolean);
}

/***** バイタル・定型文 *****/
function genVitals(){
  const rnd=(min,max)=> Math.floor(Math.random()*(max-min+1))+min;
  const sys=rnd(110,150), dia=rnd(70,90), bpm=rnd(60,90), spo=rnd(93,99);
  const tmp=(Math.round((Math.random()*(36.9-35.8)+35.8)*10)/10).toFixed(1);
  return `vital ${sys}/${dia}/${bpm}bpm / SpO2:${spo}%  ${tmp}℃`;
}
function getPresets(){
  ensureAuxSheets_();
  const s = sh('定型文'); const lr = s.getLastRow();
  if (lr < 2) {
    return [
      {cat:'所見',label:'特記事項なし',text:'特記事項なし。経過良好。'},
      {cat:'所見',label:'バイタル安定',text:'バイタル安定。生活指導継続。'},
      {cat:'所見',label:'請求書・領収書受渡',text:'請求書・領収書を受け渡し済み。'},
      {cat:'所見',label:'配布物受渡',text:'配布物（説明資料）を受け渡し済み。'},
      {cat:'所見',label:'再同意書受渡',text:'再同意書を受け渡し済み。通院予定の確認をお願いします。'},
      {cat:'所見',label:'再同意取得確認',text:'再同意の取得を確認。引き続き施術を継続。'}
    ];
  }
  const vals = s.getRange(2,1,lr-1,3).getDisplayValues(); // [カテゴリ, ラベル, 文章]
  return vals.map(r=>({cat:r[0],label:r[1],text:r[2]}));
}

/***** 施術保存 *****/
function queueAfterTreatmentJob(job){
  const p = PropertiesService.getScriptProperties();
  const key = 'AFTER_JOBS';
  const jobs = JSON.parse(p.getProperty(key) || '[]');
  jobs.push(job);
  p.setProperty(key, JSON.stringify(jobs));

  // 1分後に afterTreatmentJob を実行
  ScriptApp.newTrigger('afterTreatmentJob')
    .timeBased().after(1000 * 60).create();
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
  const p = PropertiesService.getScriptProperties();
  const key = 'AFTER_JOBS';
  const jobs = JSON.parse(p.getProperty(key) || '[]');
  p.deleteProperty(key);
  if (!jobs.length) return;

  jobs.forEach(job=>{
    const pid = job.patientId;

    // News / 同意日 / 負担割合 / 予定登録など重い処理をここでまとめて実行
    if (job.presetLabel){
      if (job.presetLabel.indexOf('再同意取得確認') >= 0){
        const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'Asia/Tokyo','yyyy-MM-dd');
        updateConsentDate(pid, today);
      }
      if (job.presetLabel.indexOf('再同意書受渡') >= 0){
        pushNews_(pid,'再同意','再同意書を受け渡し');
      }
    }
    if (job.burdenShare){
      updateBurdenShare(pid, job.burdenShare);
    }
    if (job.visitPlanDate){
      sh('予定').appendRow([pid,'通院', job.visitPlanDate, (Session.getActiveUser()||{}).getEmail()]);
      pushNews_(pid,'予定','通院予定を登録：' + job.visitPlanDate);
    }
    log_('施術後処理', pid, JSON.stringify(job));
  });
}


/***** 当月の施術一覧 取得・更新・削除 *****/
function listTreatmentsForCurrentMonth(pid){
  const s=sh('施術録'); const lr=s.getLastRow(); if(lr<2) return [];
  const vals=s.getRange(2,1,lr-1,6).getValues(); // A..F
  const tz=Session.getScriptTimeZone()||'Asia/Tokyo';
  const now=new Date();
  const start=new Date(now.getFullYear(), now.getMonth(), 1, 0,0,0);
  const end  =new Date(now.getFullYear(), now.getMonth()+1, 0, 23,59,59);

  const out=[];
  for(let i=0;i<vals.length;i++){
    const r=vals[i]; const ts=r[0]; const id=String(r[1]);
    if(id!==String(pid)) continue;
    const d = ts instanceof Date ? ts : new Date(ts);
    if(isNaN(d.getTime())) continue;
    out.push({
      row: 2+i,
      when: Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm'),
      note: String(r[2]||''),
      email: String(r[3]||'')
    });
  }
  return out.reverse();
}
function updateTreatmentRow(row, note) {
  const s = sh('施術録');
  if (row <= 1 || row > s.getLastRow()) throw new Error('行が不正です');

  const newNote = String(note || '').trim();

  // 直前の値を取得
  const oldNote = String(s.getRange(row, 3).getValue() || '').trim();

  // 🔒 二重編集チェック
  if (oldNote === newNote) {
    return { ok: false, skipped: true, msg: '変更内容が直前と同じのため編集をスキップしました' };
  }

  // 書き換え
  s.getRange(row, 3).setValue(newNote);

  // ログ
  log_('施術修正', '(row:' + row + ')', newNote);

  return { ok: true, updatedRow: row, newNote };
}

function deleteTreatmentRow(row){
  const s=sh('施術録'); if(row<=1 || row>s.getLastRow()) throw new Error('行が不正です');
  s.deleteRow(row);
  log_('施術削除', '(row:'+row+')', '');
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
function updateConsentDate(pid, dateStr){
  const hit = findPatientRow_(pid);
  if (!hit) throw new Error('患者が見つかりません');
  const s=sh('患者情報'); const head=hit.head;
  const cCons= getColFlexible_(head, LABELS.consent, PATIENT_COLS_FIXED.consent, '同意年月日');
  s.getRange(hit.row, cCons).setValue(dateStr);
  pushNews_(pid,'同意','再同意取得確認（同意日更新：'+dateStr+'）');
  clearConsentRelatedNews_(pid);
  log_('同意日更新', pid, dateStr);
}
function updateBurdenShare(pid, shareText){
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
  pushNews_(pid,'通知','負担割合を更新：' + disp);
  log_('負担割合更新', pid, disp);

  // 4) 施術録にも記録を残す（監査・検索用）
  const user = (Session.getActiveUser()||{}).getEmail();
  sh('施術録').appendRow([new Date(), String(pid), '負担割合を更新：' + (disp || shareText || ''), user, '', '' ]);

  return true;
}


/***** 請求集計（回数/負担） *****/
function rebuildInvoiceForMonth_(year, month){
  const ssb = ss();
  const t = sh('施術録'); const p = sh('患者情報');
  const outName = year + '年' + month + '月分';
  let out = ssb.getSheetByName(outName); if(!out) out = ssb.insertSheet(outName); else out.clear();
  out.getRange(1,1,1,4).setValues([['施術録番号','患者様氏名','合計施術回数','負担割合']]);

  const plc = p.getLastColumn(), plr=p.getLastRow();
  const ph = p.getRange(1,1,1,plc).getDisplayValues()[0];
  const cRec = resolveColByLabels_(ph, LABELS.recNo, '施術録番号');
  const cName= resolveColByLabels_(ph, LABELS.name,  '名前');
  const cShare=resolveColByLabels_(ph, LABELS.share, '負担割合');
  const pvals = plr>1 ? p.getRange(2,1,plr-1,plc).getDisplayValues() : [];
  const pmap = {};
  pvals.forEach(r=>{
    const rec=String(r[cRec-1]||'').trim(); if(!rec) return;
    pmap[rec]={ name:r[cName-1]||'', share:r[cShare-1]||'' };
  });

  const tlr=t.getLastRow(); if(tlr<2) return;
  const tvals=t.getRange(2,1,tlr-1,6).getValues();
  const start=new Date(year, month-1, 1, 0,0,0);
  const end  =new Date(year, month, 0, 23,59,59);
  const counts={};
  tvals.forEach(r=>{
    const ts=r[0], id=String(r[1]||'').trim(); if(!id) return;
    const d = ts instanceof Date ? ts : new Date(ts); if(isNaN(d.getTime())) return;
    if(d>=start && d<=end) counts[id]=(counts[id]||0)+1;
  });

  const rows=[];
  Object.keys(counts).sort((a,b)=> (parseInt(a,10)||0)-(parseInt(b,10)||0)).forEach(rec=>{
    const info=pmap[rec]||{name:'',share:''};
    rows.push([rec, info.name, counts[rec], info.share]);
  });
  if(rows.length) out.getRange(2,1,rows.length,4).setValues(rows);
}
function rebuildInvoiceForCurrentMonth(){
  const now=new Date(); rebuildInvoiceForMonth_(now.getFullYear(), now.getMonth()+1);
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
function composeViaOpenAI_(type, header, notes){
  const key = getOpenAiKey_();
  if (!key) return null;
  const sys = [
    'あなたは在宅リハ・鍼灸の文書作成を支援するアシスタントです。',
    '与えられたメモ（バイタル/痛み/経過/注意点）を、指定された提出先に相応しい口調と見出しで1ページ相当の日本語文書に整形してください。',
    '医療上の断定は避け、事実ベースで簡潔に。機微情報は含めません。'
  ].join('\n');
  const audience =
    type==='care_manager' ? 'ケアマネジャー向け：生活・動作の観点を重視。依頼事項は箇条書きで簡潔に。' :
    type==='family'       ? 'ご家族向け：専門用語は避け、分かりやすい表現。安心感のある文体。' :
                            '同意医師向け：簡潔な臨床情報（バイタル/疼痛/機能/施術反応）を記載。依頼は明確に。';
  const prompt =
`【患者概要】
氏名:${header.name||'-'}（ID:${header.patientId}）
年齢:${header.age||'-'} / 同意日:${header.consentDate||'-'} / 次回期限:${header.consentExpiry||'-'}
当月:${header.monthly.current.count}回 / 前月:${header.monthly.previous.count}回

【メモ】
- バイタル: ${notes?.vital||''}
- 痛み・動作: ${notes?.pain||''}
- 施術反応: ${notes?.response||''}
- 注意点・依頼: ${notes?.note||''}

【提出先】
${audience}

以上を踏まえ、提出先に合わせた1ページ相当の文書（見出し＋本文、丁寧語）に整形してください。`;

  const payload = {
    model: APP.OPENAI_MODEL,
    messages: [
      { role:'system', content: sys },
      { role:'user',   content: prompt }
    ],
    temperature: 0.2,
    max_tokens: 800
  };
  const res = UrlFetchApp.fetch(APP.OPENAI_ENDPOINT, {
    method: 'post',
    headers: { 'Authorization':'Bearer '+key, 'Content-Type':'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code >= 300) throw new Error('文章整形APIエラー: ' + code + ' ' + res.getContentText());
  const data = JSON.parse(res.getContentText());
  const content = data?.choices?.[0]?.message?.content || '';
  return content.trim();
}
function composeNarrativeLocal_(type, notes){
  const v = (notes && notes.vital)    ? String(notes.vital).trim()    : '';
  const p = (notes && notes.pain)     ? String(notes.pain).trim()     : '';
  const r = (notes && notes.response) ? String(notes.response).trim() : '';
  const n = (notes && notes.note)     ? String(notes.note).trim()     : '';
  const squash = (t)=> t ? t.replace(/\n+/g,' / ').replace(/・/g,'').trim() : '';
  const vital = squash(v), pain = squash(p), resp = squash(r), note = squash(n);
  if (type === 'care_manager') {
    return `【概況】${vital||'バイタルは概ね安定しています。'}\n【生活・動作】${pain||'疼痛や可動域の大きな変化は認めません。'}\n【施術経過】${resp||'施術反応は良好です。'}\n【連携のお願い】${note||'次回同意手続き・通院予定の調整につきご協力ください。'}`;
  } else if (type === 'family') {
    return `■ からだの様子：${vital||'体調はおおむね安定しています。'}\n■ 痛みや動き：${pain||'日常生活動作で大きな支障は見られません。'}\n■ 施術のようす：${resp||'施術後はすっきりされています。'}\n■ おねがい：${note||'通院の予定や書類の準備が必要な場合は、事前にお知らせいたします。'}`;
  } else {
    return `【バイタル/現症】${vital||'特記すべき変動はありません。'}\n【疼痛/機能】${pain||'疼痛の増悪なくADLは維持されています。'}\n【施術反応】${resp||'施術後の筋緊張の緩和を得ています。'}\n【所見/依頼】${note||'再同意取得のご検討と必要時のご指示をお願いします。'}`;
  }
}

function composeIcfViaOpenAI_(header, source){
  const key = getOpenAiKey_();
  if (!key) return null;

  const sys = [
    'あなたは在宅リハビリテーションの専門家であり、国際生活機能分類（ICF）の観点で要約を作成します。',
    '活動（Activity）、参加（Participation）、環境因子（Environmental factors）の3項目について、それぞれ2〜3文の日本語でまとめてください。',
    '所見は事実ベースで簡潔に。推測や断定は避け、丁寧語で記述してください。'
  ].join('\n');

  const notesText = Array.isArray(source?.treatments) && source.treatments.length
    ? source.treatments.map(n => {
        const note = n.note ? `所見:${n.note}` : '所見:（記録なし）';
        const vital = n.vitals ? `バイタル:${n.vitals}` : 'バイタル:（記録なし）';
        return `- ${n.when}: ${note} / ${vital}`;
      }).join('\n')
    : '（施術録情報なし）';

  const metricsText = Array.isArray(source?.metrics) && source.metrics.length
    ? source.metrics.map(metric => {
        const entries = metric.points.map(p => `${p.date}: ${p.value}${metric.unit || ''}${p.note ? `（${p.note}）` : ''}`);
        return `- ${metric.label}: ${entries.join(', ') || '（記録なし）'}`;
      }).join('\n')
    : '（臨床指標の記録なし）';

  const handoverText = Array.isArray(source?.handovers) && source.handovers.length
    ? source.handovers.map(h => `- ${h.when}: ${h.note || '（内容なし）'}`).join('\n')
    : '（申し送り情報なし）';

  const prompt = [
    `患者: ${header.name || '-'}（ID:${header.patientId}）`,
    `年齢: ${header.age || '-'} / 同意日:${header.consentDate || '-'} / 最近の施術回数: 当月${header.monthly?.current?.count||0}回`,
    `対象期間: ${source?.rangeLabel || '全期間'}`,
    '',
    '【最近の施術メモ】',
    notesText || '（メモ情報なし）',
    '',
    '【臨床指標】',
    metricsText || '（定量指標の記録なし）',
    '',
    '【申し送り】',
    handoverText || '（申し送り情報なし）',
    '',
    '上記を踏まえ、JSON形式で回答してください。キーは activity, participation, environment の3つです。',
    '各キーの値は丁寧語の文章（2〜3文程度）。例: {"activity":"...","participation":"...","environment":"..."}',
    '出力はJSONのみとし、他の文字列は含めないでください。'
  ].join('\n');

  const payload = {
    model: APP.OPENAI_MODEL,
    messages: [
      { role:'system', content: sys },
      { role:'user',   content: prompt }
    ],
    temperature: 0.2,
    max_tokens: 600
  };

  try {
    const res = UrlFetchApp.fetch(APP.OPENAI_ENDPOINT, {
      method: 'post',
      headers: { 'Authorization':'Bearer '+key, 'Content-Type':'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code >= 300) throw new Error('ICFサマリAPIエラー: ' + code + ' ' + res.getContentText());
    const data = JSON.parse(res.getContentText());
    const content = data?.choices?.[0]?.message?.content || '';
    if (!content) return null;
    let parsed = null;
    try {
      parsed = JSON.parse(content);
    } catch(e) {
      parsed = null;
    }
    if (!parsed) return null;
    return {
      via: 'ai',
      sections: [
        { key: 'activity', title: '活動', body: String(parsed.activity || '').trim() },
        { key: 'participation', title: '参加', body: String(parsed.participation || '').trim() },
        { key: 'environment', title: '環境因子', body: String(parsed.environment || '').trim() },
      ]
    };
  } catch (err) {
    Logger.log('composeIcfViaOpenAI_ error: ' + err);
    return null;
  }
}

function composeIcfLocal_(source){
  const generic = {
    activity: '日常生活動作は大きな変化なく、必要な支援量で対応しています。',
    participation: '家族や支援者と連携しながら、社会参加の機会を維持できています。',
    environment: '住環境や支援体制に大きな変更はなく、必要時にスタッフがフォローしています。'
  };

  const treatmentSummary = Array.isArray(source?.treatments) && source.treatments.length
    ? source.treatments.map(n => `${n.when}: ${n.note || n.raw || '所見なし'}`).join(' / ')
    : '';
  const handoverSummary = Array.isArray(source?.handovers) && source.handovers.length
    ? source.handovers.map(h => `${h.when}: ${h.note || '内容なし'}`).join(' / ')
    : '';
  const metricSummary = Array.isArray(source?.metrics) && source.metrics.length
    ? source.metrics.map(m => `${m.label}: ${m.points.map(p => `${p.date} ${p.value}${m.unit || ''}`).join(', ')}`).join(' / ')
    : '';

  const sections = [
    {
      key: 'activity',
      title: '活動',
      body: treatmentSummary || metricSummary || generic.activity
    },
    {
      key: 'participation',
      title: '参加',
      body: handoverSummary || treatmentSummary || generic.participation
    },
    {
      key: 'environment',
      title: '環境因子',
      body: metricSummary || handoverSummary || generic.environment
    }
  ].map(sec => ({ key: sec.key, title: sec.title, body: String(sec.body || generic[sec.key]).trim() }));

  return { via: 'local', sections };
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


function buildDoctorReportTemplate_(header, context, statusText, specialItems){
  const hospital = header?.hospital ? String(header.hospital).trim() : '';
  const doctor = header?.doctor ? String(header.doctor).trim() : '';
  const name = header?.name ? String(header.name).trim() : `ID:${header?.patientId || ''}`;
  const birth = header?.birth ? String(header.birth).trim() : '';
  const consent = context?.consentText ? String(context.consentText).trim() : '情報不足';
  const frequency = context?.frequencyLabel ? String(context.frequencyLabel).trim() : '情報不足';
  const status = statusText ? String(statusText).trim() : '（情報不足のため生成できません）';
  let specialLines = [];
  if (Array.isArray(specialItems)) {
    specialLines = specialItems.map(s => String(s || '').trim()).filter(Boolean);
  } else {
    const raw = String(specialItems || '').trim();
    if (raw) {
      let parsed = null;
      if (/^\[.*\]$/.test(raw)) {
        try {
          const arr = JSON.parse(raw);
          if (Array.isArray(arr)) parsed = arr;
        } catch (e) {
          parsed = null;
        }
      }
      if (Array.isArray(parsed)) {
        specialLines = parsed.map(s => String(s || '').trim()).filter(Boolean);
      } else {
        specialLines = raw
          .split(/\n+/)
          .map(s => String(s || '').trim())
          .filter(Boolean);
      }
    }
  }
  specialLines = specialLines.filter(line => line && line !== '[]').slice(0, 3);
  const special = specialLines.length
    ? specialLines.map(line => `・${line}`).join('\n')
    : '施術前にバイタル値を確認し、リスク管理を徹底しております。';
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const createdAt = Utilities.formatDate(new Date(), tz, 'yyyy年M月d日');
  return [
    `【病院名】${hospital || '不明'}`,
    `【担当医名】${doctor || '不明'}`,
    `【患者氏名】${name || '—'}`,
    `【生年月日】${birth || '不明'}`,
    `【同意内容】${consent}`,
    `【施術頻度】${frequency}`,
    '【患者の状態・経過】',
    status,
    '【特記すべき事項】',
    special,
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

function extractSpecialPointsFallback_(handovers){
  const latest = Array.isArray(handovers)
    ? handovers
        .filter(h => String(h?.note || '').trim())
        .slice(-5)
        .reverse()
    : [];
  if (!latest.length) return [];
  const keywords = ['嚥下','水分','食事','服薬','服用','検査','家族','介護','活動','参加','歩行','転倒','ストレス','睡眠','排泄','栄養','痛み','バイタル','血圧','SpO2','脈拍','体温','ADL','IADL'];
  const picked = [];
  latest.forEach(entry => {
    const sentences = String(entry.note)
      .split(/[。\n]/)
      .map(s => s.trim())
      .filter(Boolean);
    sentences.forEach(sentence => {
      if (picked.length >= 3) return;
      if (keywords.some(k => sentence.indexOf(k) >= 0)) {
        if (!picked.includes(sentence)) picked.push(sentence);
      }
    });
  });
  return picked;
}

function resolveReportTypeMeta_(reportType){
  const key = String(reportType || '').toLowerCase();
  const baseFormatInstruction = [
    'JSONのみを出力してください。',
    'キーは "status" と "special" の2つのみです。',
    '"status" は文字列で、"special" は文字列の配列で返してください。',
    '余分な文字列や注釈は含めないでください。'
  ].join('\n');
  const baseSpecialInstruction = [
    '共有したい特記事項が複数ある場合は配列の各要素として整理してください。',
    '該当事項がない場合でも空配列にはせず、["特記すべき事項はありません。"] を返してください。'
  ].join('\n');

  const map = {
    family: {
      key: 'family',
      label: '家族向け報告書',
      systemTone: 'あなたは在宅ケアスタッフです。ご家族に安心感を与える、やさしく分かりやすい丁寧語で状況を説明してください。専門用語は避け、落ち着いた口調で伝えます。',
      statusInstruction: '家族が安心できるように穏やかな丁寧語で最近の様子と今後の見守り方針をまとめ、「同意内容に沿った施術を継続しております。」という一文を必ず含めてください。',
      specialInstruction: [
        'ご家庭で共有したい配慮点や生活上の注意をICFの視点で整理してください。',
        baseSpecialInstruction
      ].join('\n'),
      specialLabel: '家族と共有したいポイント',
      formatInstruction: baseFormatInstruction
    },
    caremanager: {
      key: 'caremanager',
      label: 'ケアマネ向け報告書',
      systemTone: 'あなたは在宅ケアの専門職です。ケアマネジャーがケアプランに反映しやすいよう、客観的で事務的な口調で報告します。',
      statusInstruction: '活動・参加・環境因子の視点で経過を整理し、丁寧語で2〜3段落にまとめ、「同意内容に沿った施術を継続しております。」という一文を必ず含めてください。',
      specialInstruction: [
        '多職種連携で共有したい着眼点や支援上の留意事項を列挙してください。',
        baseSpecialInstruction
      ].join('\n'),
      specialLabel: 'ケアマネ連携ポイント',
      formatInstruction: baseFormatInstruction
    },
    doctor: {
      key: 'doctor',
      label: '医師向け報告書',
      systemTone: 'あなたは在宅リハに携わる医療専門職です。医学的な視点で経過を整理し、事実ベースの丁寧語で主治医に報告してください。',
      statusInstruction: '臨床経過と現在の評価、今後の施術方針や必要な連携事項を2〜3段落の丁寧語でまとめ、「同意内容に沿った施術を継続しております。」という一文を必ず含めてください。',
      specialInstruction: [
        '医師に共有したい特記や臨床的注意点があれば箇条書きで示してください。',
        baseSpecialInstruction
      ].join('\n'),
      specialLabel: '医師向け特記すべき事項',
      formatInstruction: baseFormatInstruction
    }
  };
  return map[key] || map.doctor;
}

function composeAiReportViaOpenAI_(header, context, reportType){
  const key = getOpenAiKey_();
  if (!key) return null;

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const rangeStart = context?.startDate instanceof Date && !isNaN(context.startDate.getTime())
    ? Utilities.formatDate(context.startDate, tz, 'yyyy-MM-dd')
    : '記録開始から';
  const rangeEnd = context?.endDate instanceof Date && !isNaN(context.endDate.getTime())
    ? Utilities.formatDate(context.endDate, tz, 'yyyy-MM-dd')
    : Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const sanitizeEntries = (entries, formatter, limit) => {
    if (!Array.isArray(entries) || !entries.length) return [];
    const mapped = [];
    entries.forEach(item => {
      const text = formatter(item);
      if (text) mapped.push(text);
    });
    return typeof limit === 'number' && limit > 0 ? mapped.slice(-limit) : mapped;
  };

  const handoverLines = sanitizeEntries(
    context?.handovers,
    h => {
      const note = String(h?.note || '').trim();
      if (!note) return '';
      return `- ${h.when || ''}: ${note}`.trim();
    },
    10
  );
  if (!handoverLines.length) handoverLines.push('- 情報なし');

  const treatmentLines = sanitizeEntries(
    context?.treatments,
    t => {
      const body = String(t?.raw || t?.note || '').trim();
      if (!body) return '';
      return `- ${t.when || ''}: ${body}`.trim();
    },
    10
  );
  if (!treatmentLines.length) treatmentLines.push('- 情報なし');

  const metricLines = [];
  if (Array.isArray(context?.metrics)) {
    context.metrics.forEach(metric => {
      const pts = Array.isArray(metric?.points) ? metric.points.slice(-5) : [];
      if (!pts.length) return;
      const entries = pts
        .map(p => {
          const value = isFinite(p?.value) ? `${p.value}${metric.unit || ''}` : '';
          const note = String(p?.note || '').trim();
          const parts = [p?.date || '', value, note ? `備考:${note}` : ''].filter(Boolean);
          return parts.length ? parts.join(' ') : '';
        })
        .filter(Boolean);
      if (entries.length) {
        metricLines.push(`- ${metric.label}: ${entries.join(', ')}`);
      }
    });
  }
  if (!metricLines.length) metricLines.push('- 情報なし');

  const consentText = String(context?.consentText || '').trim() || '情報不足';
  const frequency = String(context?.frequencyLabel || '').trim() || '情報不足';
  const meta = resolveReportTypeMeta_(reportType);

  const sys = [
    meta.systemTone,
    '提供されたデータのみを根拠に、事実ベースで丁寧語（へりくだり口調）でまとめてください。',
    '参考情報の文章をそのままコピーせず、要点のみを抽出してください。',
    '回答はJSONのみとし、"status" と "special" の2キー以外は含めないでください。'
  ].join('\n');

  const rawData = [
    '【患者基本情報】',
    `- 氏名: ${header?.name || '-'}`,
    `- 患者ID: ${header?.patientId || ''}`,
    `- 生年月日: ${header?.birth || '-'}`,
    `- 年齢: ${header?.age != null ? `${header.age}歳${header?.ageClass ? '（' + header.ageClass + '）' : ''}` : '情報不足'}`,
    `- 同意日: ${header?.consentDate || '情報不足'}`,
    `- 同意内容: ${consentText}`,
    `- 施術頻度: ${frequency}`,
    `- 対象期間: ${context?.rangeLabel || '全期間'}（${rangeStart}〜${rangeEnd}）`,
    '',
    '【申し送り（参考情報）】',
    handoverLines.join('\n'),
    '',
    '【施術録抜粋（参考情報）】',
    treatmentLines.join('\n'),
    '',
    '【臨床指標（参考情報）】',
    metricLines.join('\n')
  ].join('\n');

  const instructionLines = [
    meta.formatInstruction,
    `- "status": ${meta.statusInstruction}`,
    `- "special": ${meta.specialInstruction}`
  ].join('\n');

  const prompt = [
    '以下の要件に従ってAI報告書をJSONで出力してください。',
    '',
    instructionLines,
    '',
    '注意: 参考情報コピー禁止、要点のみ利用。',
    '',
    '出力例:',
    '{',
    '  "status": "ここに本文",',
    '  "special": ["ここに特記事項1", "ここに特記事項2"]',
    '}',
    '',
    '参考情報:',
    rawData
  ].join('\n');

  try {
    const res = UrlFetchApp.fetch(APP.OPENAI_ENDPOINT, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + key, 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        model: APP.OPENAI_MODEL,
        messages: [
          { role: 'system', content: sys },
          { role: 'user', content: prompt }
        ],
        temperature: 0.2,
        max_tokens: 900
      }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code >= 300) {
      throw new Error('AI報告書APIエラー: ' + code + ' ' + res.getContentText());
    }
    const content = (JSON.parse(res.getContentText())?.choices?.[0]?.message?.content || '').trim();
    if (!content) return null;

    let plain = content.trim();
    if (!plain) return null;

    if (/^```/.test(plain)) {
      plain = plain.replace(/^```[^\n]*\n?/, '');
    }
    if (/```$/.test(plain)) {
      plain = plain.replace(/```$/, '');
    }

    const firstBrace = plain.indexOf('{');
    const lastBrace = plain.lastIndexOf('}');
    if (firstBrace >= 0 && lastBrace >= firstBrace) {
      plain = plain.slice(firstBrace, lastBrace + 1);
    }

    let parsed = null;
    try {
      parsed = JSON.parse(plain);
    } catch (err) {
      Logger.log('composeAiReportViaOpenAI_ JSON parse error: ' + err);
      return null;
    }
    if (!parsed || typeof parsed !== 'object') return null;

    let status = String(parsed.status || '').trim();
    if (!status) return null;
    if (status.indexOf('同意内容に沿った施術を継続しております。') < 0) {
      status += (status.endsWith('。') ? '' : '。') + '同意内容に沿った施術を継続しております。';
    }
    status = status.trim();

    let specialRaw = parsed.special;
    let special = [];
    if (Array.isArray(specialRaw)) {
      special = specialRaw;
    } else if (specialRaw != null) {
      if (typeof specialRaw === 'string') {
        const trimmed = specialRaw.trim();
        if (trimmed) {
          if (/^\[.*\]$/.test(trimmed)) {
            try {
              const arr = JSON.parse(trimmed);
              if (Array.isArray(arr)) special = arr;
            } catch (e) {
              special = [];
            }
          }
          if (!special.length) {
            special = trimmed
              .split(/\n+/)
              .map(s => s.replace(/^[-*・]\s*/, '').trim())
              .filter(Boolean);
          }
        }
      }
    }
    special = Array.isArray(special) ? special.map(item => String(item || '').trim()).filter(Boolean) : [];
    if (!special.length) {
      special = ['特記すべき事項はありません。'];
    }

    return {
      via: 'ai',
      status,
      special
    };
  } catch (err) {
    Logger.log('composeAiReportViaOpenAI_ error: ' + err);
    return null;
  }
}

function composeAiReportLocal_(header, context, reportType){
  const meta = resolveReportTypeMeta_(reportType);
  const name = header?.name || `ID:${header?.patientId || ''}`;
  const rangeLabel = context?.rangeLabel || '直近の期間';
  const handovers = Array.isArray(context?.handovers) ? context.handovers : [];
  const metrics = Array.isArray(context?.metrics) ? context.metrics : [];
  const freqLabel = String(context?.frequencyLabel || '').trim();
  const consentText = String(context?.consentText || '').trim();

  const latestHandover = handovers.filter(h => String(h?.note || '').trim()).slice(-1)[0];
  const metricDigest = buildMetricDigestForSummary_(metrics) || '';
  const handoverDigestFamily = buildHandoverDigestForSummary_(handovers, 'family') || '';
  const handoverDigestCare = buildHandoverDigestForSummary_(handovers, 'caremanager') || '';
  const doctorDigest = buildHandoverDigestForSummary_(handovers, 'doctor') || '';

  const statusParagraphs = [];
  if (meta.key === 'family') {
    statusParagraphs.push(`${name}様の${rangeLabel}のご様子についてご報告申し上げます。`);
    if (handoverDigestFamily) {
      statusParagraphs.push(handoverDigestFamily.replace(/最近のようす：/, '最近のようすでは').replace(/。$/, '。'));
    } else if (latestHandover) {
      statusParagraphs.push(`最近は「${latestHandover.note}」との申し送りがあり、落ち着いた経過で推移されています。`);
    } else {
      statusParagraphs.push('大きな変化は確認されておらず、落ち着いた状態で過ごされています。');
    }
    const third = [];
    if (metricDigest) {
      third.push(`臨床指標では ${metricDigest} が確認されています。`);
    }
    third.push('今後も生活のリズムを崩さないよう見守りながら施術を進めてまいります。');
    statusParagraphs.push(third.join(' '));
  } else if (meta.key === 'caremanager') {
    statusParagraphs.push(`${rangeLabel}の状況をICFの視点でご報告いたします。`);
    statusParagraphs.push(handoverDigestCare || '活動・参加面では大きな変化は見られず、現状維持で経過しています。');
    const envParts = [];
    if (metricDigest) {
      envParts.push(`臨床指標: ${metricDigest}。`);
    } else {
      envParts.push('臨床指標: 直近の数値記録は確認されませんでした。');
    }
    envParts.push(`環境因子: ${consentText || '既存の支援体制を継続しています。'}`);
    envParts.push(`施術頻度: ${freqLabel || '情報不足（直近の件数が不足）'}`);
    statusParagraphs.push(envParts.join(' '));
  } else {
    if (doctorDigest) {
      statusParagraphs.push(doctorDigest.replace(/最近の申し送りでは、/, '最近の申し送りでは').replace(/。$/, '。'));
    } else {
      statusParagraphs.push('申し送りの記録からは急激な変化は確認されておりません。');
    }
    const secondParagraph = [];
    if (freqLabel) secondParagraph.push(`直近1か月の施術頻度は${freqLabel}です。`);
    if (metricDigest) secondParagraph.push(`臨床指標では ${metricDigest} が確認されています。`);
    secondParagraph.push('今後も必要に応じて評価を続けてまいります。');
    statusParagraphs.push(secondParagraph.join(' '));
  }

  if (statusParagraphs.length) {
    const lastIndex = statusParagraphs.length - 1;
    if (statusParagraphs[lastIndex].indexOf('同意内容に沿った施術を継続しております。') < 0) {
      statusParagraphs[lastIndex] += (statusParagraphs[lastIndex].endsWith('。') ? '' : '。') + '同意内容に沿った施術を継続しております。';
    }
  } else {
    statusParagraphs.push('同意内容に沿った施術を継続しております。');
  }

  const status = statusParagraphs.filter(Boolean).join('\n\n').trim();

  let special = extractSpecialPointsFallback_(handovers) || [];
  if (!special.length && metricDigest) {
    special = [metricDigest];
  }
  if (!special.length && freqLabel) {
    special = [`施術頻度: ${freqLabel}`];
  }
  if (!special.length) {
    special = ['特記すべき事項はありません。'];
  }

  return {
    via: 'local',
    status,
    special
  };
}

function normalizeReportSpecial_(special){
  const arr = Array.isArray(special) ? special : [];
  const normalized = arr.map(item => String(item || '').trim()).filter(Boolean);
  return normalized.length ? normalized : ['特記すべき事項はありません。'];
}

function generateAiReport(payload){
  assertDomain_();
  ensureAuxSheets_();
  const pidRaw = payload && typeof payload === 'object' ? (payload.patientId || payload.pid || payload.id) : payload;
  const pid = String(pidRaw || '').trim();
  if (!pid) throw new Error('患者IDを指定してください');

  const rangeKey = payload && typeof payload === 'object' && payload.range ? String(payload.range) : '1m';
  const reportTypeRaw = payload && typeof payload === 'object' && payload.reportType ? String(payload.reportType) : '';
  const reportMeta = resolveReportTypeMeta_(reportTypeRaw);
  const range = resolveIcfSummaryRange_(rangeKey);
  const header = getPatientHeader(pid);
  if (!header) throw new Error('患者が見つかりません');

  const treatments = getTreatmentNotesInRange_(pid, range.startDate, range.endDate);
  const handovers = getHandoversInRange_(pid, range.startDate, range.endDate);
  const metrics = listClinicalMetricSeries(pid, range.startDate, range.endDate).metrics;
  const consentText = getConsentContentForPatient_(pid);
  const treatmentCount = countTreatmentsInRecentMonth_(pid, range.endDate);
  const frequencyLabel = determineTreatmentFrequencyLabel_(treatmentCount);
  const metricCount = metrics.reduce((sum, metric) => sum + (Array.isArray(metric.points) ? metric.points.length : 0), 0);

  const context = {
    rangeKey,
    rangeLabel: range.label,
    startDate: range.startDate,
    endDate: range.endDate,
    treatments,
    handovers,
    metrics,
    consentText,
    frequencyLabel,
    treatmentCount
  };

  let report = composeAiReportViaOpenAI_(header, context, reportMeta.key);
  if (!report) {
    report = composeAiReportLocal_(header, context, reportMeta.key);
  }
  if (!report || !report.status) {
    throw new Error('AI報告書を生成できませんでした');
  }

  const specialNormalized = normalizeReportSpecial_(report.special);
  report.special = specialNormalized;

  const sheet = ensureAiReportSheet_();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const now = new Date();
  sheet.appendRow([
    now,
    header.patientId,
    range.label,
    reportMeta.label,
    report.status || '',
    JSON.stringify(specialNormalized)
  ]);

  log_('AI報告書生成', header.patientId, `${range.label}/${reportMeta.label}`);

  const generatedAt = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm');
  return {
    ok: true,
    usedAi: report.via === 'ai',
    rangeKey,
    rangeLabel: range.label,
    generatedAt,
    reportType: reportMeta.key,
    audienceLabel: reportMeta.label,
    specialLabel: reportMeta.specialLabel,
    meta: {
      handoverCount: handovers.length,
      metricCount,
      treatmentCount,
      frequencyLabel,
      patientFound: true
    },
    report: {
      status: report.status || '',
      special: specialNormalized
    }
  };
}

/***** レポートPDF（API→フォールバック） *****/
function getPatientHeaderForReport_(pid){
  const header = getPatientHeader(pid);
  if (!header) throw new Error('患者が見つかりません');
  return header;
}
function generateReportViaApi(pid, type, notes){
  assertDomain_(); ensureAuxSheets_();
  const header = getPatientHeaderForReport_(pid);

  let narrative = null;
  try { narrative = composeViaOpenAI_(type, header, notes); } catch(e){ narrative = null; }
  if (!narrative) narrative = composeNarrativeLocal_(type, notes);

  const tz = Session.getScriptTimeZone()||'Asia/Tokyo';
  const title = `report_${type}_${Utilities.formatDate(new Date(), tz,'yyyyMM')}.pdf`;
  const body =
`[提出先:${type}]
${header.name||''} 様  /  ID:${header.patientId}
年齢:${header.age||'-'} (${header.ageClass||''})
同意日:${header.consentDate||'-'} / 次回期限:${header.consentExpiry||'-'}
当月:${header.monthly.current.count}回, 前月:${header.monthly.previous.count}回

${narrative}
`;
  return savePdf_(pid, title, body);
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
  SpreadsheetApp.getUi()
    .createMenu('請求')
    .addItem('今月の集計（回数+負担割合）','rebuildInvoiceForCurrentMonth')
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
    if (touched.size) DashboardIndex_updatePatients(Array.from(touched));
    // キャッシュは雑に全無効化（運用後にキー粒度を最適化）
    CacheService.getScriptCache().removeAll();
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

  // 入力（例: "2025-09-04T14:30" / "2025-09-04 14:30" / "2025/9/4 14:30"）を Date に変換
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const d = parseDateTimeFlexible_(newLocal, tz);
  if (!d || isNaN(d.getTime())) throw new Error('日時の形式が不正です');

  // 書き換え
  s.getRange(row, 1).setValue(d);

  // 監査ログ
  const toDisp = (v)=> v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm') : String(v||'');
  log_('施術TS修正', pid, `row=${row}  ${toDisp(oldTs)} -> ${toDisp(d)}`);
  pushNews_(pid, '記録', `施術記録の日時を修正: ${toDisp(d)}`);

  // ダッシュボードの最終施術日に影響するので Index を更新（v1は全件でOK）
  DashboardIndex_updatePatients([pid]);

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
function submitTreatment(payload) {
  try {
    ensureAuxSheets_();
    const s = sh('施術録');
    const pid = String(payload?.patientId || '').trim();
    if (!pid) throw new Error('patientIdが空です');

    const user = (Session.getActiveUser() || {}).getEmail() || '';

    // タイムゾーンを日本時間に固定して文字列保存
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

    const vit = (payload?.overrideVitals || '').trim() || genVitals();
    const note = String(payload?.notesParts?.note || '').trim();
    const merged = note ? (note + '\n' + vit) : vit;

    // 🔒 二重保存チェック（直近の1件と比較）
    const lr = s.getLastRow();
    if (lr >= 2) {
      const last = s.getRange(lr, 1, 1, 4).getValues()[0]; // [TS, pid, 所見, user]
      const lastPid = String(last[1] || '').trim();
      const lastNote = String(last[2] || '').trim();
      if (lastPid === pid && lastNote === merged) {
        return { ok: false, skipped: true, msg: '直前と同じ内容のため保存をスキップしました' };
      }
    }

    const row = [now, pid, merged, user, '', ''];
    s.appendRow(row);

    if (Array.isArray(payload?.clinicalMetrics) && payload.clinicalMetrics.length) {
      recordClinicalMetrics_(pid, payload.clinicalMetrics, now, user);
    }

    return { ok: true, vitals: vit, wroteTo: s.getName(), row };
  } catch (e) {
    throw e;
  }
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
