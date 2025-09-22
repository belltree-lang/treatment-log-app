/***** スケジュール管理用シートID *****/
const SCHEDULE_SS_ID = "1Xq39zEUVHoCRTOE4vQpPsgoPMhXsuHQAPTG5RriNY2o";

/**
 * 指定 leadId に対して TrialOpportunities を生成
 */
function generateTrialOpportunities(leadId) {
  Logger.log("受け取った leadId: " + leadId);
  const now = new Date();
  const cfg = getConfig() || {};
  Logger.log("cfg読み込み: " + JSON.stringify(cfg));

  const lookAheadDays = parseInt(cfg.look_ahead_days || 10, 10);
  const serviceMin = parseInt(cfg.service_duration_min || 45, 10);
  const tz = cfg.timezone || "Asia/Tokyo";

  // --- Waitlistから対象患者取得 ---
  const schedSS = SpreadsheetApp.openById(SCHEDULE_SS_ID);
  const wlSheet = schedSS.getSheetByName("Waitlist");
  const wlValues = wlSheet.getDataRange().getValues();
  const wlHeaders = wlValues[0];
  const wlIdx = wlHeaders.indexOf("leadId");
  if (wlIdx === -1) throw new Error("WaitlistにleadId列がありません");

  const wlRow = wlValues.find(r => r[wlIdx] === leadId);
  if (!wlRow) throw new Error("指定LeadIdが見つかりません: " + leadId);

  const wlObj = {};
  wlHeaders.forEach((h,i)=> wlObj[h] = wlRow[i]);

  const patientName = wlObj.name;
  const patientAddr = wlObj.address;
  const patientLoc  = geocodeAddress(patientAddr);
  if (!patientLoc) throw new Error("住所をジオコーディングできません: " + patientAddr);

  // --- スケジュール系シート ---
  const slotSheet = schedSS.getSheetByName("SlotMaster");
  const assignSheet = schedSS.getSheetByName("Assignment");
  const trialSheet = schedSS.getSheetByName("TrialOpportunities");

  // --- SlotMaster ---
  const slots = slotSheet.getDataRange().getValues();
  const slotHeaders = slots[0];
  const slotList = slots.slice(1).map(r=>{
    const obj = {};
    slotHeaders.forEach((h,i)=>obj[h]=r[i]);
    if (obj.duration_min) obj.duration_min = parseInt(obj.duration_min,10);
    return obj;
  }).filter(s=> String(s.active).toUpperCase()==="TRUE");

  // --- Assignment ---
  const assigns = assignSheet.getDataRange().getValues();
  const assignHeaders = assigns[0];
  const assignList = assigns.slice(1).map(r=>{
    const obj={};
    assignHeaders.forEach((h,i)=>obj[h]=r[i]);
    return obj;
  }).filter(a=> a.status==="assigned");

  // --- TrialOpportunities 初期化 ---
  const existingRows = trialSheet.getLastRow();
  if(existingRows>1) trialSheet.deleteRows(2, existingRows-1);

  const results = [];

  // --- 直近N日分をチェック ---
  for(let d=0; d<lookAheadDays; d++){
    const date = new Date(now.getFullYear(), now.getMonth(), now.getDate()+d);
    const weekday = weekdayStrFromDate(date);
    Logger.log(`[DAY] ${Utilities.formatDate(date, tz, 'yyyy-MM-dd')} getDay=${date.getDay()} -> ${weekday}`);

    slotList.forEach(slot=>{
      if(slot.weekday!==weekday){
        Logger.log(`skip slot=${slot.slot_id} reason=weekday_mismatch slot.weekday=${slot.weekday} need=${weekday}`);
        return;
      }

      const slotStartStr = String(slot.start_time || "");
      if (!/^\d{1,2}:\d{2}$/.test(slotStartStr)) {
        Logger.log(`SlotMaster invalid time=${slotStartStr}`);
        return;
      }
      const [h,m] = slotStartStr.split(":").map(Number);
      const slotStart = new Date(date.getFullYear(), date.getMonth(), date.getDate(), h, m);
      const slotEnd = new Date(slotStart.getTime() + slot.duration_min*60000);

      // すでにアサイン済み？
      const assigned = assignList.find(a => {
        if (a.slot_id !== slot.slot_id) return false;
        return isActiveOn(date, a.effective_from, a.effective_to);
      });
      if(assigned){
        Logger.log(`drop slot=${slot.slot_id} reason=already_assigned`);
        return;
      }

      // 移動時間チェック
      const travelPrev = travelMinutes(patientLoc, patientLoc);
      const travelNext = travelMinutes(patientLoc, patientLoc);
      const totalRequired = travelPrev + serviceMin + travelNext;
      const slack = slot.duration_min - totalRequired;

      if(slack<0){
        Logger.log(`drop slot=${slot.slot_id} reason=insufficient_slack dur=${slot.duration_min} req=${totalRequired}`);
        return;
      }

      results.push({
        date: Utilities.formatDate(date,tz,"yyyy-MM-dd"),
        staff: slot.staff,
        area: slot.area || "",
        gap_start: slotStart,
        gap_end: slotEnd,
        gap_minutes: slot.duration_min,
        candidate_patient_id: leadId,
        candidate_name: patientName,
        candidate_address: patientAddr,
        travel_prev_min: travelPrev,
        service_min: serviceMin,
        travel_next_min: travelNext,
        total_required_min: totalRequired,
        slack_min: slack,
        reason: "OK",
        note: "SlotMaster候補"
      });
    });
  }

  // --- TrialOpportunities に書き込み ---
  if(results.length){
    const headers = trialSheet.getDataRange().getValues()[0];
    const rows = results.map(r=> headers.map(h=> r[h]||""));
    trialSheet.getRange(trialSheet.getLastRow()+1,1,rows.length,headers.length).setValues(rows);
  }

  // ログ出力
  Logger.log("=== Trial結果 件数: " + results.length + " ===");
  results.forEach(r => {
    Logger.log(`候補: date=${r.date}, staff=${r.staff}, start=${r.gap_start}, end=${r.gap_end}, slack=${r.slack_min}`);
  });

  // UI用に返す
  return results.map(r => ({
    staff: r.staff,
    area: r.area || "",
    start: Utilities.formatDate(r.gap_start, tz, "yyyy-MM-dd HH:mm"),
    end:   Utilities.formatDate(r.gap_end, tz, "yyyy-MM-dd HH:mm"),
    leadId: leadId
  }));
}


/**
 * Configシートから値を取得
 */
function getConfig(){
  const schedSS = SpreadsheetApp.openById(SCHEDULE_SS_ID);
  const cfgSheet = schedSS.getSheetByName("Config");
  if(!cfgSheet) return {};
  const values = cfgSheet.getDataRange().getValues();
  const map = {};
  values.forEach(r=>{
    if(r[0] && !String(r[0]).startsWith("#")) map[r[0]] = r[1];
  });
  return map;
}

/** 曜日文字列 */
function weekdayStrFromDate(d){
  return ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][d.getDay()];
}

/** 日付セル→Date（空ならnull） */
function toDateOrNull(v){
  if (!v) return null;
  if (v instanceof Date) return v;
  const s = String(v).trim();
  const m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

/** 指定日が [from,to] の範囲に含まれるか */
function isActiveOn(date, from, to){
  const f = toDateOrNull(from);
  const t = toDateOrNull(to);
  const d0 = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  if (f && d0 < new Date(f.getFullYear(), f.getMonth(), f.getDate())) return false;
  if (t && d0 > new Date(t.getFullYear(), t.getMonth(), t.getDate())) return false;
  return true;
}

/***** ユーティリティ（Geo/Travel） *****/
function geocodeAddress(addr){
  if(!addr) return null;
  const schedSS = SpreadsheetApp.openById(SCHEDULE_SS_ID);
  const S = schedSS.getSheetByName('GeoCache');
  const values = S.getDataRange().getValues();
  const head = values[0]||['address','lat','lng'];
  if(values.length===0) S.getRange(1,1,1,3).setValues([head]);
  const map = new Map(values.slice(1).map(r=>[String(r[0]), {lat:r[1], lng:r[2]}]));
  if(map.has(addr)) return map.get(addr);
  const res = Maps.newGeocoder().setLanguage('ja').geocode(addr);
  const r = res?.results?.[0]; if(!r) return null;
  const loc = r.geometry.location;
  S.appendRow([addr,loc.lat,loc.lng]);
  return {lat:loc.lat,lng:loc.lng};
}

/** 移動時間: Configに force_zero_travel=TRUE があれば常に0 */
function travelMinutes(origin, destination){
  const cfg = getConfig() || {};
  if (String(cfg.force_zero_travel).toUpperCase() === 'TRUE') return 0;
  if (!origin || !destination) return 0;
  return 10; // ダミー値
}

/** Assignmentを一気に空にするテスト用 */
function clearAssignmentForTest(){
  const ss = SpreadsheetApp.openById(SCHEDULE_SS_ID);
  const sh = ss.getSheetByName('Assignment');
  if (!sh) throw new Error('Assignmentシートが見つかりません');
  const last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last-1);
  Logger.log('Assignment rows cleared (kept header).');
}

/** SlotMasterを静的チェック */
function validateSlotMaster(){
  const schedSS = SpreadsheetApp.openById(SCHEDULE_SS_ID);
  const slotSheet = schedSS.getSheetByName("SlotMaster");
  const values = slotSheet.getDataRange().getValues();
  const head = values[0];
  const rows = values.slice(1);

  const validW = new Set(['Sun','Mon','Tue','Wed','Thu','Fri','Sat']);
  rows.forEach((r,i)=>{
    const rowNo = i+2;
    const rec = {}; head.forEach((h,idx)=> rec[h]=r[idx]);
    if (!validW.has(String(rec.weekday))) Logger.log(`SlotMaster row${rowNo} invalid weekday=${rec.weekday}`);
    const t = String(rec.start_time||'');
    if (!/^\d{1,2}:\d{2}$/.test(t)) Logger.log(`SlotMaster row${rowNo} invalid time=${t}`);
    if (String(rec.active).toUpperCase()!=='TRUE') Logger.log(`SlotMaster row${rowNo} inactive`);
  });
}
/**
 * Waitlist一覧
 */
function waitlistList() {
  const schedSS = SpreadsheetApp.openById(SCHEDULE_SS_ID);
  const sheet = schedSS.getSheetByName('Waitlist');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  Logger.log("Waitlist headers: " + JSON.stringify(headers));

  if (values.length <= 1) return [];

  return values.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, idx) => obj[h.trim()] = r[idx]);
    Logger.log("Row obj: " + JSON.stringify(obj));
    return {
      leadId: obj.leadId || '',
      ts: obj.ts ? String(obj.ts) : '',
      name: obj.name || '',
      phone: String(obj.phone || ''),
      address: obj.address || '',
      avoidSlotsJson: obj.avoidSlotsJson || '',
      complaint: obj.complaint || '',
      notes: obj.notes || '',
      status: obj.status || 'open',
      source: obj.source || ''
    };
  });
}
function validateSlotMaster(){
  const schedSS = SpreadsheetApp.openById(SCHEDULE_SS_ID);
  const slotSheet = schedSS.getSheetByName("SlotMaster");
  const values = slotSheet.getDataRange().getValues();
  const head = values[0];
  const rows = values.slice(1);

  const validW = new Set(['Sun','Mon','Tue','Wed','Thu','Fri','Sat']);
  rows.forEach((r,i)=>{
    const rowNo = i+2;
    const rec = {}; head.forEach((h,idx)=> rec[h]=r[idx]);
    if (!validW.has(String(rec.weekday))) Logger.log(`SlotMaster row${rowNo} invalid weekday=${rec.weekday}`);
    const t = String(rec.start_time||'');
    if (!/^\d{1,2}:\d{2}$/.test(t)) Logger.log(`SlotMaster row${rowNo} invalid time=${t}`);
    if (String(rec.active).toUpperCase()!=='TRUE') Logger.log(`SlotMaster row${rowNo} inactive`);
    if (!validW.has(String(rec.weekday))) {
  Logger.log(`SlotMaster row${rowNo} invalid weekday=${rec.weekday}`);
} else {
  Logger.log(`SlotMaster row${rowNo} weekday=${rec.weekday} OK`);
}
  });
}

/**
 * SlotMasterにテスト用の正常データを流し込む
 * （既存の内容は消えるので注意！）
 */
function setupTestSlotMaster(){
  const ss = SpreadsheetApp.openById(SCHEDULE_SS_ID);
  const sh = ss.getSheetByName('SlotMaster');
  if (!sh) throw new Error("SlotMasterシートが見つかりません");

  // ヘッダー行
  const headers = [
    "slot_id","staff","weekday","start_time","duration_min","visit_type","area","active","priority"
  ];
  
  const data = [
    // slot_id            staff   weekday start  dur visit area   active priority
    ["TST-Mon-0900-45", "テスト太郎", "Mon", "09:00", 45, "在宅", "町田", true, 1],
    ["TST-Mon-1000-45", "テスト太郎", "Mon", "10:00", 45, "在宅", "町田", true, 1],
    ["TST-Tue-1300-45", "テスト次郎", "Tue", "13:00", 45, "在宅", "八王子", true, 1],
    ["TST-Wed-1500-45", "テスト花子", "Wed", "15:00", 45, "在宅", "日野", true, 1],
    ["TST-Thu-0930-45", "テスト太郎", "Thu", "09:30", 45, "在宅", "町田", true, 1],
    ["TST-Fri-1600-45", "テスト次郎", "Fri", "16:00", 45, "在宅", "八王子", true, 1],
    ["TST-Sat-1000-45", "テスト花子", "Sat", "10:00", 45, "在宅", "日野", true, 1],
  ];

  sh.clear(); // 既存データ削除
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.getRange(2,1,data.length,headers.length).setValues(data);

  Logger.log("SlotMasterをテストデータで初期化しました。");
}
