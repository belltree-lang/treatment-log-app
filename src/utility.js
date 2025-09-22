/***** 共通ユーティリティ（utility.gs） *****/

/** スプレッドシート本体 */
function ss(){
  // ここで直接固定（プロパティよりも優先される）
  const FIXED_SSID = '1ajnW9Fuvu0YzUUkfTmw0CrbhrM3lM5tt5OA1dK2_CoQ';
  return SpreadsheetApp.openById(FIXED_SSID);
}

/** 既存互換エイリアス（schedule 側の SS() 呼び出しを壊さないため） */
function SS(){ return ss(); }

/** シート取得（存在しなければエラー） */
function sh(name){
  const wb = ss();
  const s = wb.getSheetByName(name);
  if (!s) throw new Error('シートが見つかりません: ' + name);
  return s;
}

/** プロパティ取得（schedule 側の getConfig 互換） */
function getConfig(key){
  return PropertiesService.getScriptProperties().getProperty(key);
}

/** 日付フォーマット */
function fmtDate(d, tz){
  return Utilities.formatDate(d, tz || getConfig('timezone') || 'Asia/Tokyo', 'yyyy-MM-dd');
}
function fmtDT(d, tz){
  return Utilities.formatDate(d, tz || getConfig('timezone') || 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");
}

/** 時間・丸め・曜日ユーティリティ */
function parseHHMM(s){
  const [H, M] = String(s||'').split(':').map(Number);
  return (isFinite(H)?H:0)*60 + (isFinite(M)?M:0);
}
function ceilToGrid(dateObj, minutesPerGrid){
  const d = new Date(dateObj);
  const m = d.getMinutes();
  const r = m % minutesPerGrid;
  if (r !== 0){
    d.setMinutes(m + (minutesPerGrid - r));
    d.setSeconds(0);
    d.setMilliseconds(0);
  }
  return d;
}
function weekdayAbbr(d){ return ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][d.getDay()]; }
function weekdayToByDay(wk){ return {Sun:'SU',Mon:'MO',Tue:'TU',Wed:'WE',Thu:'TH',Fri:'FR',Sat:'SA'}[wk]; }

/** いま参照している書き込み先を丸見えにする */
function debugEnv(){
  const wb = ss();
  return {
    TARGET_ID: wb.getId(),
    TARGET_URL: wb.getUrl(),
    SHEETS: wb.getSheets().map(s=>s.getName())
  };
}

function debugCheck() {
  const s = sh('施術録');
  Logger.log('対象スプレッドシートURL: ' + ss().getUrl());
  Logger.log('シート名: ' + s.getName());
  Logger.log('最終行: ' + s.getLastRow());
  if (s.getLastRow() >= 2) {
    Logger.log('直近データ: ' + JSON.stringify(s.getRange(s.getLastRow(),1,1,6).getValues()));
  }
}

/** SSID を ScriptProperties に記録（初回セット用） */
function setSpreadsheetId(id) {
  if (!id) throw new Error('IDが空です');
  PropertiesService.getScriptProperties().setProperty('SSID', String(id));
}

function initSetId() {
  setSpreadsheetId('1ajnW9Fuvu0YzUUkfTmw0CrbhrM3lM5tt5OA1dK2_CoQ');
}