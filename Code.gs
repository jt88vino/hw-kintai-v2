// 連携先スプレッドシートのID
const SPREADSHEET_ID = "1SjRkvg9kk1YFjJKROBlxXFHp10xO3_mC-JSq9yAM9lk";

// マスターシート名
const MASTER_SHEET_NAME = "勤怠マスタ";
const MEMBER_SHEET_SUFFIX = "_勤務表";
const CATEGORY_USERS = ["田中", "牛嶋", "長谷川", "住吉", "鈴木"];
const WORK_CATEGORIES = ["アカデミー", "ホームワイン", "その他（WT業務）"];
// シフト（予定表）: 種別ごとに1行。削除は行を消さず「状態」に残す（新着一覧のため）。
const SHIFT_SHEET_NAME = "シフト";
const SHIFT_HEADERS = ["ID","種別","名前","開始日","終了日","開始時刻","終了時刻","ラベル","メモ","登録者","登録日時","状態","削除者","削除日時"];
const SHIFT_KINDS = ["シフト","業務","出勤","休み","有給","振休"];
// お知らせ: 管理者（アプリの管理者認証、または NOTICE_TOKEN プロパティ）だけが掲載・削除できる。
const NOTICE_SHEET_NAME = "お知らせ";
const NOTICE_HEADERS = ["ID","本文","掲載開始","掲載終了","登録者","登録日時","状態","削除日時"];
const ALLOCATION_ITEMS = [{"id": "hw_production", "label": "HWの生産", "description": "伝票作成/詰め替え/梱包（WT）", "category": "ホームワイン"}, {"id": "hw_support", "label": "HWのお問い合わせ", "description": "メール/Lステップ（HW）", "category": "ホームワイン"}, {"id": "hw_admin", "label": "HWの管理", "description": "発送完了メール/売上処理/ec force配送管理/搬入（HW）", "category": "ホームワイン"}, {"id": "hw_pro", "label": "HWのPRO制作", "description": "PROのH1/Figma（HW）", "category": "ホームワイン"}, {"id": "hw_other", "label": "HWのその他", "description": "搬入などボトル販売/イベント/ツアー ※ホームワイン人件費に含まれない項目", "category": "その他（WT業務）"}, {"id": "hwa_production", "label": "HWAの生産", "description": "伝票作成/詰め替え/梱包", "category": "アカデミー"}, {"id": "hwa_support", "label": "HWAのお問い合わせ", "description": "メール/Lステップ", "category": "アカデミー"}, {"id": "hwa_admin", "label": "HWAの管理", "description": "発送完了メール/売上処理/ec force配送管理/搬入", "category": "アカデミー"}, {"id": "hwa_text", "label": "HWAのテキスト制作", "description": "H1/Figma（HW）、キャンバ、動画", "category": "アカデミー"}, {"id": "hwa_other", "label": "HWAのその他", "description": "搬入などボトル販売/イベント/ツアー ※アカデミー人件費に含まれない項目", "category": "その他（WT業務）"}];
// 管理者認証情報はスクリプトプロパティで管理する。

// マスターシートのヘッダー
const MASTER_HEADERS = [
  "タイムスタンプ",
  "名前",
  "区分",
  "打刻日時",
  "対象月",
  "システムID",
  "交通機関",
  "備考",
  "業務区分",
  "実働時間",
  ...ALLOCATION_ITEMS.map(item=>item.label),
  "業務配分データ"
];

/**
 * GETリクエスト受信用
 * HTML側からの action=read / action=add / action=delete /
 * action=addUser / action=deleteUser をここで処理します。
 */
function doGet(e) {
  return handleRequest_(e, false);
}

/**
 * POSTリクエスト受信用
 * 念のためPOSTでも同じ処理を通します。
 */
function doPost(e) {
  return handleRequest_(e, true);
}

/**
 * メイン処理
 */
function handleRequest_(e, isPost) {
  e = e || {};
  const params = e.parameter || {};
  const action = params.action || "read";

  const adminActions = ["adminLogin", "addUser", "deleteUser", "delete"];
  if (action === "adminResult") {
    const key = String(params.receipt || "");
    const value = /^[a-f0-9-]{36}$/.test(key) ? CacheService.getScriptCache().get("result:" + key) : null;
    return createResponse_(e, value ? JSON.parse(value) : {ok:false, pending:true});
  }
  if (adminActions.indexOf(action) !== -1 || action === "deleteRecent" || action === "shiftAdd" || action === "shiftDelete" || action === "noticeAdd" || action === "noticeDelete" || (action === "add" && isPost)) {
    if (!isPost || params.callback) return createResponse_(e, {ok:false, error:"post_required"});
    if (!/^[a-f0-9-]{36}$/.test(String(params.receipt || ""))) return createResponse_(e, {ok:false, error:"invalid_receipt"});
    let result;
    try {
      if (action === "adminLogin") result = loginAdmin_(params);
      else if (action === "add") { const addSS = SpreadsheetApp.openById(SPREADSHEET_ID); result = addLog_(getOrCreateMasterSheet_(addSS), params); }
      else if (action === "deleteRecent") { const recentSS = SpreadsheetApp.openById(SPREADSHEET_ID); result = deleteLog_(getOrCreateMasterSheet_(recentSS), params, true); }
      else if (action === "shiftAdd") { result = addShift_(SpreadsheetApp.openById(SPREADSHEET_ID), params); }
      else if (action === "shiftDelete") { result = deleteShift_(SpreadsheetApp.openById(SPREADSHEET_ID), params); }
      else if (action === "noticeAdd") { result = addNotice_(SpreadsheetApp.openById(SPREADSHEET_ID), params); }
      else if (action === "noticeDelete") { result = deleteNotice_(SpreadsheetApp.openById(SPREADSHEET_ID), params); }
      else {
        if(action==='delete' && params.adminPassword) {
          const auth=loginAdmin_(params);
          if(auth.ok) params.adminToken=auth.token;
          else result=auth;
        }
        if(!result) result = validateAdmin_(params, action);
        if (!result) {
          const adminSS = SpreadsheetApp.openById(SPREADSHEET_ID);
          if (action === "addUser") result = addUser_(adminSS, params);
          else if (action === "deleteUser") result = deleteUser_(adminSS, params);
          else result = deleteLog_(getOrCreateMasterSheet_(adminSS), params);
        }
      }
    } catch (err) { result = {ok:false, error:"admin_action_failed", message:"管理者操作に失敗しました。再同期して状態を確認してください。"}; }
    CacheService.getScriptCache().put("result:" + params.receipt, JSON.stringify(result), 60);
    return createResponse_(e, result);
  }
  try {
    if(action==='read' && params.scope==='recent' && params.fresh!=='1') {
      const cached=CacheService.getScriptCache().get('attendance:recent');
      if(cached) return createResponse_(e,JSON.parse(cached));
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (action === "shift") return createResponse_(e, readShifts_(ss, params));
    const sheet = action === "read" ? ss.getSheetByName(MASTER_SHEET_NAME) : getOrCreateMasterSheet_(ss);

    if (action === "read") {
      return createResponse_(e, readData_(ss, sheet, params));
    }

    if (action === "add") {
      return createResponse_(e, addLog_(sheet, params));
    }

    if (action === "delete") {
      return createResponse_(e, deleteLog_(sheet, params));
    }

    if (action === "addUser") {
      return createResponse_(e, addUser_(ss, params));
    }

    if (action === "deleteUser") {
      return createResponse_(e, deleteUser_(ss, params));
    }

    return createResponse_(e, {
      ok: false,
      error: "unknown_action",
      message: "不明なactionです: " + action
    });

  } catch (err) {
    return createResponse_(e, {
      ok: false,
      error: String(err && err.message ? err.message : err),
      stack: String(err && err.stack ? err.stack : "")
    });
  }
}

/**
 * 勤怠マスタシートを取得。なければ作成。
 */
function getOrCreateMasterSheet_(ss) {
  let sheet = ss.getSheetByName(MASTER_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(MASTER_SHEET_NAME);
    sheet.appendRow(MASTER_HEADERS);
    sheet.getRange(1, 1, 1, MASTER_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(MASTER_HEADERS);
    sheet.getRange(1, 1, 1, MASTER_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    return sheet;
  }

  const headerRange = sheet.getRange(1, 1, 1, MASTER_HEADERS.length);
  const currentHeaders = headerRange.getDisplayValues()[0];
  let needsHeaderFix = false;

  for (let i = 0; i < MASTER_HEADERS.length; i++) {
    if (!currentHeaders[i]) {
      needsHeaderFix = true;
      break;
    }
  }

  if (needsHeaderFix) {
    headerRange.setValues([MASTER_HEADERS]);
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 打刻データ追加
 */
function addLog_(sheet, params) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const name = params.name || "";
    const type = params.type || "";
    const time = params.time || "";
    const month = params.month || deriveMonth_(time);
    const transport = params.transport || "";
    let memo = String(params.memo || "").split(/[\r\n｜]+/).map(s=>s.trim()).filter(Boolean).join("｜");
    const category = String(params.category || "");
    if (category && (!CATEGORY_USERS.includes(name) || !WORK_CATEGORIES.concat(["業務配分"]).includes(category))) {
      return {ok:false, action:"add", error:"invalid_category", message:"業務区分を確認してください。"};
    }

    if (!name || !type || !time) {
      return {
        ok: false,
        action: "add",
        error: "missing_required_params",
        message: "name / type / time のいずれかが不足しています。"
      };
    }

    // HTML側が発行したIDをそのまま保存する。同一リクエストが再送されても、
    // 同じIDまたは同じ打刻内容なら二重登録しない。
    const logId = String(params.requestId || params.id || new Date().getTime());
    const existingLogId = findExistingLogId_(sheet, logId, name, type, time, category);
    if (existingLogId) {
      return {
        ok: true,
        action: "add",
        id: existingLogId,
        duplicate: true,
        message: "すでに登録済みのため、重複登録を防止しました。"
      };
    }

    let allocations = [];
    if (params.allocations) {
      try { allocations = JSON.parse(params.allocations); } catch(e) { return {ok:false,error:"invalid_allocation",message:"業務時間の形式を確認してください。"}; }
    }
    if (category === "業務配分" && type === "退勤") {
      const error = validateAllocations_(allocations);
      if (error) return error;
      const readIndex=getReadIndex_(sheet);
      const history=readIndexedRows_(sheet,(readIndex.sessions[name]||[]).map(n=>[n,n]));
      const actual = shiftMinutes_(history, name, time);
      if (actual === null) return {ok:false,error:"missing_clockin",message:"出勤の記録が見つかりません。同期して確認してください。"};
      if (allocations.reduce((n,a)=>n+a.minutes,0)!==actual) return {ok:false,error:"allocation_mismatch",actualMinutes:actual,message:"配分合計を実働"+actual+"分に合わせてください。"};
      const details=allocations.filter(a=>a.minutes>0 || a.memo).map(a=>{
        const item=ALLOCATION_ITEMS.find(i=>i.id===a.id);
        return item.label+"（"+a.minutes+"分）"+(a.memo?"："+String(a.memo).split(/[\r\n｜]+/).map(s=>s.trim()).filter(Boolean).join("｜"):"");
      });
      if(memo) details.push(memo);
      memo=details.join("｜");
    } else if (allocations.length) return {ok:false,error:"unexpected_allocation",message:"業務配分は対象者の退勤時に入力してください。"};

    CacheService.getScriptCache().remove('attendance:recent');
    sheet.appendRow([
      new Date(),
      name,
      type,
      time,
      month,
      logId,
      transport,
      memo,
      category,
      "",
      ...ALLOCATION_ITEMS.map(item=>{const a=allocations.find(a=>a.id===item.id);return a?a.minutes/1440:"";}),
      allocations.length?JSON.stringify(allocations):""
    ]);
    const rowNumber = sheet.getLastRow();
    // 同一人物・業務区分の直前の打刻が勤務中なら、その区間の実働を加算。
    // セル参照なので、管理者が打刻時刻を修正した場合にも再計算される。
    if (category && rowNumber > 2) {
      const r = rowNumber, p = r - 1;
      const match = '($B$2:$B$'+p+'=B'+r+')' + (category==='業務配分'?'':'*($I$2:$I$'+p+'=I'+r+')');
      const prevType = 'LOOKUP(2,ARRAYFORMULA(1/('+match+')),$C$2:$C$'+p+')';
      const prevTime = 'LOOKUP(2,ARRAYFORMULA(1/('+match+')),$D$2:$D$'+p+')';
      sheet.getRange(r,10).setFormula('=IFERROR(IF(AND(OR(C'+r+'="退勤",C'+r+'="休憩開始"),OR('+prevType+'="出勤",'+prevType+'="休憩終了"),VALUE(D'+r+')-VALUE('+prevTime+')<=0.75),MAX(0,VALUE(D'+r+')-VALUE('+prevTime+')),0),0)').setNumberFormat('[h]:mm:ss');
    }
    if(allocations.length) sheet.getRange(rowNumber,11,1,10).setNumberFormat("[h]:mm");
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove('attendance:recent');

    return {
      memo: memo,
      allocations: allocations,
      ok: true,
      action: "add",
      id: logId,
      message: "SUCCESS"
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * 同じ送信ID、または名前・区分・打刻日時が完全一致する行を探す。
 * 通信遅延による同一URLの再送を安全に1件へまとめる。
 */
function findExistingLogId_(sheet, logId, name, type, time, category) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "";

  // Search only ID/time columns on the Sheets side; fetch full rows only for candidates.
  const idMatch=sheet.getRange(2,6,lastRow-1,1).createTextFinder(String(logId)).matchEntireCell(true).findNext();
  if(idMatch) return String(logId);
  const candidates=sheet.getRange(2,4,lastRow-1,1).createTextFinder(String(time)).matchEntireCell(true).findAll();
  for(const candidate of candidates) {
    const row=sheet.getRange(candidate.getRow(),1,1,9).getDisplayValues()[0];
    if(row[1]===name && row[2]===type && row[3]===time && (row[8]||'')===(category||'')) return String(row[5]||logId);
  }

  return "";
}

/**
 * データ読み取り
 */
// The index contains row locations only; attendance remains in the master sheet.
// Chunk properties to stay below the per-value limit. Reads and writes share a lock.
const READ_INDEX_PREFIX = 'ATTENDANCE_READ_INDEX_';
function invalidateAttendanceIndex_() {
  PropertiesService.getScriptProperties().deleteProperty(READ_INDEX_PREFIX + 'meta');
  CacheService.getScriptCache().remove('attendance:recent');
}
function attendanceSheetChanged(e) {
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    CacheService.getScriptCache().remove('attendance:roster');
    if (!e || !e.range || e.range.getSheet().getName() === MASTER_SHEET_NAME) invalidateAttendanceIndex_();
  } finally {lock.releaseLock();}
}
function setupAttendanceReadIndex() {
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t=>t.getHandlerFunction()==='attendanceSheetChanged')) {
    ScriptApp.newTrigger('attendanceSheetChanged').forSpreadsheet(SPREADSHEET_ID).onChange().create();
  }
  invalidateAttendanceIndex_();
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  readData_(ss,ss.getSheetByName(MASTER_SHEET_NAME),{scope:'recent'});
}
function loadReadIndex_() {
  const props=PropertiesService.getScriptProperties();
  try {
    const meta=JSON.parse(props.getProperty(READ_INDEX_PREFIX+'meta')||'null');
    if(!meta) return null;
    let value='';
    for(let i=0;i<meta.parts;i++) value+=props.getProperty(READ_INDEX_PREFIX+i)||'';
    return JSON.parse(value);
  } catch(e) { return null; }
}
function saveReadIndex_(index) {
  const props=PropertiesService.getScriptProperties(), value=JSON.stringify(index);
  const parts=Math.ceil(value.length/2000); // Japanese text can take three UTF-8 bytes.
  const old=JSON.parse(props.getProperty(READ_INDEX_PREFIX+'meta')||'null');
  for(let i=0;i<parts;i++) props.setProperty(READ_INDEX_PREFIX+i,value.slice(i*2000,(i+1)*2000));
  for(let i=parts;old && i<old.parts;i++) props.deleteProperty(READ_INDEX_PREFIX+i);
  props.setProperty(READ_INDEX_PREFIX+'meta',JSON.stringify({parts:parts}));
}
function indexTime_(value) {
  const text=String(value||'').replace(/\//g,'-');
  const m=text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return m ? m[1]+'-'+m[2].padStart(2,'0')+'-'+m[3].padStart(2,'0')+' '+m[4].padStart(2,'0')+':'+m[5]+':'+(m[6]||'00') : text;
}
function logFromRow_(row, rowNumber) {
  return {id:String(row[5]||rowNumber-1),name:row[1]||'',type:row[2]||'',time:row[3]||'',
    month:row[4]||deriveMonth_(row[3]),transport:row[6]||'',memo:row[7]||'',
    category:row[8]||'',allocations:parseAllocations_(row[20])};
}
function updateReadIndex_(index, rows, startRow) {
  const entries=[];
  rows.forEach((row,i)=>{
    if(!row[1] && !row[2] && !row[3]) return;
    const n=startRow+i, month=String(row[4]||deriveMonth_(row[3]));
    const bucket=index.months[month]||(index.months[month]={spans:[],last:{}});
    const tail=bucket.spans[bucket.spans.length-1];
    if(tail && tail[1]===n-1) tail[1]=n; else bucket.spans.push([n,n]);
    const time=indexTime_(row[3]);
    entries.push({n:n,name:row[1],type:row[2],time:time,month:month,category:row[8]||''});
  });
  index.sessionStart=index.sessionStart||{};
  const timeOf=t=>new Date(String(t).replace(' ','T')+'+09:00').getTime();
  entries.sort((a,b)=>a.time.localeCompare(b.time)||a.n-b.n).forEach(e=>{
    const key=JSON.stringify([e.name,e.category]);
    index.months[e.month].last[key]=[e.n,e.time];
    index.latest[e.name]=e.n;
    let session=index.sessions[e.name]||[];
    // An open session past the shift limit is over; the next punch starts from nothing.
    const started=index.sessionStart[e.name];
    if(session.length && started && timeOf(e.time)-timeOf(started)>SHIFT_LIMIT_MS) session=[];
    if(e.type==='出勤' && !session.length){session=[e.n]; index.sessionStart[e.name]=e.time;}
    else if(e.type==='退勤') session=[];
    else if(session.length) session.push(e.n);
    index.sessions[e.name]=session;
    index.recent.push(e.n); if(index.recent.length>50) index.recent.shift();
    index.maxTime=e.time;
  });
  index.lastRow=startRow+rows.length-1;
  index.revision=Utilities.getUuid();
}
function getReadIndex_(sheet) {
  const lastRow=sheet?sheet.getLastRow():0;
  let index=loadReadIndex_();
  // A rebuild also reconciles changes made by external APIs (which do not fire edit triggers).
  if(!index || index.version!==1 || index.lastRow>lastRow || Date.now()-index.builtAt>21600000) index=null;
  if(index && index.lastRow<lastRow) {
    const rows=sheet.getRange(index.lastRow+1,1,lastRow-index.lastRow,MASTER_HEADERS.length).getDisplayValues();
    if(rows.some(r=>r[3] && indexTime_(r[3])<index.maxTime)) index=null;
    else {updateReadIndex_(index,rows,index.lastRow+1);saveReadIndex_(index);}
  }
  if(!index) {
    index={version:1,lastRow:lastRow,builtAt:Date.now(),months:{},latest:{},sessions:{},recent:[],maxTime:'',revision:Utilities.getUuid()};
    if(lastRow>1) updateReadIndex_(index,sheet.getRange(2,1,lastRow-1,MASTER_HEADERS.length).getDisplayValues(),2);
    saveReadIndex_(index);
  }
  return index;
}
function readIndexedRows_(sheet, spans) {
  if(!sheet || !spans.length) return [];
  // Merge overlaps and adjacent ranges; never expand gaps into a full-history read.
  const merged=[];
  spans.sort((a,b)=>a[0]-b[0]).forEach(span=>{
    const tail=merged[merged.length-1];
    if(tail && span[0]<=tail[1]+1) tail[1]=Math.max(tail[1],span[1]);
    else merged.push(span.slice());
  });
  const logs=[];
  merged.forEach(span=>sheet.getRange(span[0],1,span[1]-span[0]+1,MASTER_HEADERS.length).getDisplayValues().forEach((r,i)=>{
    if(r[1] || r[2] || r[3]) logs.push(logFromRow_(r,span[0]+i));
  }));
  return logs.sort((a,b)=>indexTime_(b.time).localeCompare(indexTime_(a.time)));
}
function readData_(ss, sheet, params) {
  params=params||{};
  const scope=params.scope||'legacy', month=String(params.month||'');
  if(!['recent','month','legacy'].includes(scope)) return {ok:false,error:'invalid_scope'};
  if(scope==='month' && !/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) return {ok:false,error:'invalid_month'};
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if(scope==='recent' && params.fresh!=='1') {
      const cached=CacheService.getScriptCache().get('attendance:recent');
      if(cached) return JSON.parse(cached);
    }
    const index=getReadIndex_(sheet);
    let spans=[];
    if(scope==='month') {
      spans=(index.months[month]?index.months[month].spans:[]).map(s=>s.slice());
      // Previous category punch is context for a legacy interval crossing a month boundary.
      const previous={};
      Object.keys(index.months).filter(m=>m<month).sort().forEach(m=>{
        Object.entries(index.months[m].last).forEach(([key,value])=>{
          if(!previous[key] || previous[key][1]<value[1]) previous[key]=value;
        });
      });
      Object.values(previous).forEach(v=>spans.push([v[0],v[0]]));
    } else if(scope==='legacy') {
      // Compatibility for app tabs opened before deployment. New clients always specify scope.
      if(index.lastRow>1) spans.push([2,index.lastRow]);
    } else {
      const rows=new Set(index.recent.concat(Object.values(index.latest),...Object.values(index.sessions)));
      rows.forEach(n=>spans.push([n,n]));
    }
    const logs=readIndexedRows_(sheet,spans), roster=readRoster_(ss);
    const result={ok:true,action:'read',schemaVersion:5,scope:scope,month:scope==='month'?month:null,
      revision:index.revision,months:Object.keys(index.months).sort().reverse(),logs:logs,
      users:roster.users,transportationCosts:roster.transportationCosts,notices:readNotices_(ss),
      serverTime:Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd HH:mm:ss')};
    if(scope==='recent') {
      const value=JSON.stringify(result);
      if(value.length<25000) CacheService.getScriptCache().put('attendance:recent',value,15);
    }
    return result;
  } finally {lock.releaseLock();}
}

/**
 * 「〇〇_勤務表」シートを全端末共通の名簿として読み取る。
 */
function readRoster_(ss) {
  const cache=CacheService.getScriptCache(), cached=cache.get('attendance:roster');
  if(cached) return JSON.parse(cached);
  const users = [];
  const transportationCosts = {};

  ss.getSheets().forEach(function(sheet) {
    const sheetName = sheet.getName();

    if (!sheetName.endsWith(MEMBER_SHEET_SUFFIX)) {
      return;
    }

    const name = sheetName.slice(0, -MEMBER_SHEET_SUFFIX.length);
    if (!name) {
      return;
    }

    users.push(name);

    try {
      const costValue = sheet.getRange("I1").getValue();
      const cost = parseInt(costValue, 10);
      transportationCosts[name] = isNaN(cost) ? 0 : cost;
    } catch (err) {
      transportationCosts[name] = 0;
    }
  });

  const roster={users:users,transportationCosts:transportationCosts};
  cache.put('attendance:roster',JSON.stringify(roster),300);
  return roster;
}

/**
 * 管理者画面からメンバーを追加する。
 * 既存の勤務表をテンプレートとして複製し、C1へ名前、I1へ交通費0円を設定する。
 */
function addUser_(ss, params) {
  const authError = validateAdmin_(params, "addUser");
  if (authError) return authError;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const nameResult = validateUserName_(params.name);
    if (!nameResult.ok) return nameResult;

    const name = nameResult.name;
    const newSheetName = name + MEMBER_SHEET_SUFFIX;

    if (ss.getSheetByName(newSheetName)) {
      return {
        ok: false,
        action: "addUser",
        error: "already_exists",
        message: name + "さんはすでに名簿に登録されています。"
      };
    }

    const memberSheets = getMemberSheets_(ss);
    if (!memberSheets.length) {
      return {
        ok: false,
        action: "addUser",
        error: "template_not_found",
        message: "複製元になる勤務表が見つかりません。"
      };
    }

    const newSheet = memberSheets[0].copyTo(ss).setName(newSheetName);
    newSheet.getRange("C1").setValue(name);
    newSheet.getRange("I1").setValue(0);
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(ss.getNumSheets());
    SpreadsheetApp.flush();

    CacheService.getScriptCache().remove('attendance:roster');
    CacheService.getScriptCache().remove('attendance:recent');
    const roster = readRoster_(ss);
    return {
      ok: true,
      action: "addUser",
      users: roster.users,
      transportationCosts: roster.transportationCosts,
      message: name + "さんを名簿に追加しました。"
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * 管理者画面からメンバーを削除する。
 * 個別勤務表は非表示の削除済みシートとして退避し、勤怠マスタの過去ログも残す。
 */
function deleteUser_(ss, params) {
  const authError = validateAdmin_(params, "deleteUser");
  if (authError) return authError;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const nameResult = validateUserName_(params.name);
    if (!nameResult.ok) return nameResult;

    const name = nameResult.name;
    const targetSheet = ss.getSheetByName(name + MEMBER_SHEET_SUFFIX);

    if (!targetSheet) {
      return {
        ok: false,
        action: "deleteUser",
        error: "not_found",
        message: name + "さんは名簿に登録されていません。"
      };
    }

    const memberSheets = getMemberSheets_(ss);
    if (memberSheets.length <= 1) {
      return {
        ok: false,
        action: "deleteUser",
        error: "last_user",
        message: "名簿には1人以上必要です。"
      };
    }

    preserveMonthSourceBeforeDelete_(memberSheets, targetSheet);
    const archiveName = createArchivedSheetName_(ss, name);
    targetSheet.setName(archiveName);
    targetSheet.hideSheet();
    SpreadsheetApp.flush();

    CacheService.getScriptCache().remove('attendance:roster');
    CacheService.getScriptCache().remove('attendance:recent');
    const roster = readRoster_(ss);
    return {
      ok: true,
      action: "deleteUser",
      users: roster.users,
      transportationCosts: roster.transportationCosts,
      message: name + "さんを名簿から削除しました。個別勤務表と過去ログは退避・保持されています。"
    };

  } finally {
    lock.releaseLock();
  }
}

function createArchivedSheetName_(ss, name) {
  const timestamp = Utilities.formatDate(
    new Date(),
    "Asia/Tokyo",
    "yyyyMMdd_HHmmss"
  );
  const baseName = name + MEMBER_SHEET_SUFFIX + "_削除済_" + timestamp;
  let candidate = baseName;
  let suffix = 2;

  while (ss.getSheetByName(candidate)) {
    candidate = baseName + "_" + suffix;
    suffix++;
  }

  return candidate.slice(0, 100);
}

function getMemberSheets_(ss) {
  return ss.getSheets().filter(function(sheet) {
    return sheet.getName().endsWith(MEMBER_SHEET_SUFFIX);
  });
}

/**
 * 各勤務表のB1が削除対象シートを参照している場合、別の勤務表へ参照を付け替える。
 */
function preserveMonthSourceBeforeDelete_(memberSheets, targetSheet) {
  const targetSheetName = targetSheet.getName();
  const replacementSheet = memberSheets.find(function(sheet) {
    return sheet.getSheetId() !== targetSheet.getSheetId();
  });

  if (!replacementSheet) return;

  const targetMonth = targetSheet.getRange("B1").getValue();
  const replacementMonthCell = replacementSheet.getRange("B1");
  const replacementFormula = replacementMonthCell.getFormula();

  if (replacementFormula.indexOf("'" + targetSheetName.replace(/'/g, "''") + "'") !== -1) {
    replacementMonthCell.setValue(targetMonth);
  }

  const replacementNameEscaped = replacementSheet.getName().replace(/'/g, "''");
  memberSheets.forEach(function(sheet) {
    if (sheet.getSheetId() === targetSheet.getSheetId() ||
        sheet.getSheetId() === replacementSheet.getSheetId()) {
      return;
    }

    const monthCell = sheet.getRange("B1");
    const formula = monthCell.getFormula();
    if (formula.indexOf("'" + targetSheetName.replace(/'/g, "''") + "'") !== -1) {
      monthCell.setFormula("='" + replacementNameEscaped + "'!B1");
    }
  });
}

function loginAdmin_(params) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty("ADMIN_PASSWORD_SHA256");
  if (!expected) return {ok:false, error:"admin_not_configured", message:"管理者認証が未設定です。"};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache();
    const failures = Number(cache.get("adminFailures") || 0);
    if (failures >= 20) return {ok:false, error:"rate_limited", message:"しばらく待ってから認証してください。"};
    const actual = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(params.adminPassword || ""), Utilities.Charset.UTF_8)
      .map(function(b){ return (b & 255).toString(16).padStart(2, "0"); }).join("");
    if (actual !== expected) {
      cache.put("adminFailures", String(failures + 1), 300);
      return {ok:false, error:"unauthorized", message:"パスワードが正しくありません。"};
    }
    cache.remove("adminFailures");
    const token = Utilities.getUuid() + Utilities.getUuid();
    cache.put("admin:" + token, "valid", 1800);
    return {ok:true, action:"adminLogin", token:token};
  } finally { lock.releaseLock(); }
}

function validateAdmin_(params, action) {
  const token = String(params.adminToken || "");
  if (/^[a-f0-9-]{72}$/.test(token) && CacheService.getScriptCache().get("admin:" + token) === "valid") return null;
  return {ok:false, action:action, error:"unauthorized", message:"管理者認証が必要です。もう一度ログインしてください。"};
}

function validateUserName_(value) {
  const name = String(value || "").trim().replace(/\s+/g, " ");

  if (!name) {
    return {
      ok: false,
      error: "missing_name",
      message: "追加・削除する方のお名前を入力してください。"
    };
  }

  if (name.length > 20) {
    return {
      ok: false,
      error: "name_too_long",
      message: "お名前は20文字以内で入力してください。"
    };
  }

  if (/[\\\/?*\[\]:]/.test(name)) {
    return {
      ok: false,
      error: "invalid_name",
      message: "お名前に使用できない記号が含まれています。"
    };
  }

  return { ok: true, name: name };
}

/**
 * 打刻データ削除
 * idがあればid優先。idがない場合は name + time で削除。
 */
function deleteLog_(sheet, params, recentOnly) {
  // Only the dedicated POST action may cancel the eight logs shown on the punch screen.
  const authError = recentOnly === true ? null : validateAdmin_(params, "delete");
  if (authError) return authError;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const targetId = params.id || "";
    const targetName = params.name || "";
    const targetTime = params.time || "";

    if (!targetId && (!targetName || !targetTime)) {
      return {
        ok: false,
        action: "delete",
        error: "missing_delete_params",
        message: "削除対象の id または name/time が不足しています。"
      };
    }

    if(recentOnly === true) {
      if(!targetId) return {ok:false,error:'missing_delete_params',message:'削除対象のIDが必要です。'};
      const index=getReadIndex_(sheet);
      const recent=readIndexedRows_(sheet,index.recent.map(n=>[n,n])).slice(0,8);
      if(!recent.some(log=>log.id===String(targetId))) return {
        ok:false,error:'not_recent',message:'直近の打刻ログの対象外です。再読み込みして確認してください。過去のログは勤務履歴から取り消せます。'
      };
    }
    const lastRow=sheet.getLastRow();
    let rowNumber=0;
    if(lastRow>1 && targetId) {
      const found=sheet.getRange(2,6,lastRow-1,1).createTextFinder(String(targetId)).matchEntireCell(true).matchCase(true).findNext();
      if(found) rowNumber=found.getRow();
    } else if(lastRow>1) {
      const matches=sheet.getRange(2,4,lastRow-1,1).createTextFinder(String(targetTime)).matchEntireCell(true).findAll();
      for(const found of matches.reverse()) {
        const row=sheet.getRange(found.getRow(),2,1,3).getDisplayValues()[0];
        if(row[0]===targetName && row[2]===targetTime){rowNumber=found.getRow();break;}
      }
    }
    if(rowNumber) {
      sheet.deleteRow(rowNumber);
      invalidateAttendanceIndex_();
      SpreadsheetApp.flush();
      return {ok:true,action:'delete',id:String(targetId),message:'SUCCESS'};
    }

    return {
      ok: false,
      action: "delete",
      error: "not_found",
      message: "ERROR: 対象データなし"
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * JSON / JSONP 両対応のレスポンス作成
 */
function createResponse_(e, result) {
  const params = e && e.parameter ? e.parameter : {};
  const output = JSON.stringify(result);

  if (params.callback) {
    const callback = String(params.callback);
    if (/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService
        .createTextOutput(callback + "(" + output + ");")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
  }

  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * timeから yyyy-MM を推定
 */
function deriveMonth_(time) {
  if (!time) {
    return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM");
  }

  const s = String(time);
  if (/^\d{4}-\d{2}/.test(s)) {
    return s.slice(0, 7);
  }

  if (/^\d{4}\/\d{2}/.test(s)) {
    return s.slice(0, 7).replace("/", "-");
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM");
  }

  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM");
}

function testRead() {
  const response = handleRequest_({ parameter: { action: "read" } });
  Logger.log(response.getContent());
}

function testReadJsonp() {
  const response = handleRequest_({
    parameter: { action: "read", callback: "test" }
  });
  Logger.log(response.getContent());
}

function parseAllocations_(value) {
  try { const data=JSON.parse(value||"[]"); return Array.isArray(data)?data:[]; } catch(e) { return []; }
}
function validateAllocations_(allocations) {
  if(!Array.isArray(allocations) || allocations.length!==ALLOCATION_ITEMS.length) return {ok:false,error:"invalid_allocation",message:"10項目の業務時間を確認してください。"};
  const ids=new Set();
  for(const a of allocations){
    if(!a || !ALLOCATION_ITEMS.some(i=>i.id===a.id) || ids.has(a.id) || !Number.isInteger(a.minutes) || a.minutes<0 || a.minutes>10080 || typeof a.memo!=="string" || a.memo.length>300 || (a.memo.trim() && !a.minutes)) return {ok:false,error:"invalid_allocation",message:"業務時間は0以上の分数で入力し、業務内容を記入した項目には時間も入力してください。"};
    ids.add(a.id);
  }
  return null;
}
// A shift is over 18 hours after its clock-in, whatever comes next: a forgotten clock-out
// must not pair with the next day's punches, and a later clock-in starts a new shift.
const SHIFT_LIMIT_MS = 18*60*60*1000;
function shiftMinutes_(logs,name,endTime) {
  // Sheets display values drop the leading zero on the hour ("2026-09-18 9:22:45"),
  // which is not valid ISO 8601. indexTime_ pads it before parsing.
  const parse=t=>new Date(indexTime_(t).replace(" ","T")+"+09:00").getTime();
  const end=parse(endTime); let active=false, working=false, start=0, ms=0, shiftStart=0;
  logs.filter(l=>l.name===name && parse(l.time)<=end).sort((a,b)=>parse(a.time)-parse(b.time)).forEach(l=>{
    const t=parse(l.time);
    if(active && t-shiftStart>SHIFT_LIMIT_MS){active=false;working=false;ms=0;}
    if(l.type==='出勤' && !active){active=true;working=true;start=t;shiftStart=t;ms=0;}
    else if(l.type==='休憩開始' && active && working){ms+=Math.max(0,t-start);working=false;}
    else if(l.type==='休憩終了' && active && !working){start=t;working=true;}
    else if(l.type==='退勤'){active=false;working=false;ms=0;}
  });
  if(!active || !Number.isFinite(end) || end-shiftStart>SHIFT_LIMIT_MS) return null;
  return Math.floor((ms+(working?Math.max(0,end-start):0))/60000);
}

// ===== シフト（予定表） =====
function getOrCreateShiftSheet_(ss) {
  let sheet = ss.getSheetByName(SHIFT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHIFT_SHEET_NAME);
    sheet.appendRow(SHIFT_HEADERS);
    sheet.getRange(1, 1, 1, SHIFT_HEADERS.length).setFontWeight("bold");
    sheet.getRange("A:N").setNumberFormat("@"); // keep dates and clocks as the text the app wrote
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function shiftDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Tokyo", "yyyy-MM-dd");
  const m = String(v || "").match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  return m ? m[1] + "-" + m[2].padStart(2, "0") + "-" + m[3].padStart(2, "0") : "";
}
function shiftClock_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Tokyo", "HH:mm");
  const m = String(v || "").match(/(\d{1,2}):(\d{2})/);
  return m ? m[1].padStart(2, "0") + ":" + m[2] : "";
}
function shiftEntry_(row, rowNumber) {
  return {id:String(row[0]||""), kind:String(row[1]||""), name:String(row[2]||""), start:shiftDate_(row[3]), end:shiftDate_(row[4])||shiftDate_(row[3]),
    from:shiftClock_(row[5]), to:shiftClock_(row[6]), label:String(row[7]||""), memo:String(row[8]||""), by:String(row[9]||""), at:row[10]?indexTime_(row[10]):"",
    deleted:String(row[11]||"")==="削除", deletedBy:String(row[12]||""), deletedAt:row[13]?indexTime_(row[13]):"", row:rowNumber};
}
function shiftMonths_(start, end) {
  const out = []; let y = +start.slice(0, 4), m = +start.slice(5, 7); const ey = +end.slice(0, 4), em = +end.slice(5, 7);
  while ((y < ey || (y === ey && m <= em)) && out.length < 24) { out.push(y + "-" + String(m).padStart(2, "0")); m++; if (m > 12) { m = 1; y++; } }
  return out;
}
function clearShiftCache_(start, end) { const c = CacheService.getScriptCache(); shiftMonths_(start, end).forEach(m => c.remove("shift:" + m)); }
function readShifts_(ss, params) {
  const month = String(params.month || "");
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) return {ok:false, error:"invalid_month", message:"対象月を確認してください。"};
  const cache = CacheService.getScriptCache();
  if (params.fresh !== "1") { const cached = cache.get("shift:" + month); if (cached) return JSON.parse(cached); }
  const sheet = ss.getSheetByName(SHIFT_SHEET_NAME);
  const rows = sheet && sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, SHIFT_HEADERS.length).getDisplayValues() : [];
  const first = month + "-01", last = month + "-31", entries = [], changes = [];
  rows.forEach((row, i) => {
    if (!row[0]) return;
    const e = shiftEntry_(row, i + 2);
    if (!e.deleted && e.start <= last && e.end >= first) entries.push(e);
    changes.push(e);
  });
  changes.sort((a, b) => (b.deletedAt || b.at).localeCompare(a.deletedAt || a.at));
  const result = {ok:true, action:"shift", month:month, entries:entries, changes:changes.slice(0, 30),
    serverTime:Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss")};
  const value = JSON.stringify(result);
  if (value.length < 90000) cache.put("shift:" + month, value, 60);
  return result;
}
function validateShift_(params, roster) {
  const kind = String(params.kind || ""), name = String(params.name || "").trim(), start = shiftDate_(params.start), end = shiftDate_(params.end) || start;
  const from = shiftClock_(params.from), to = shiftClock_(params.to), label = String(params.label || "").trim(), memo = String(params.memo || "").trim();
  if (SHIFT_KINDS.indexOf(kind) === -1) return {error:"invalid_kind", message:"種類を選んでください。"};
  if (!start) return {error:"invalid_date", message:"日付を確認してください。"};
  if (end < start) return {error:"invalid_date", message:"終了日は開始日以降にしてください。"};
  if (kind === "業務") { if (!label) return {error:"missing_label", message:"業務の内容を入力してください。"}; }
  else {
    if (!name) return {error:"missing_name", message:"名前を選んでください。"};
    if (roster.indexOf(name) === -1) return {error:"unknown_name", message:"名簿にない名前です。"};
    if (kind === "シフト" && end !== start) return {error:"invalid_date", message:"シフトは1日ずつ登録してください。"};
  }
  if (kind === "シフト" && (!from || !to)) return {error:"missing_time", message:"開始と終了の時刻を入力してください。"};
  if (from && to && to <= from) return {error:"invalid_time", message:"終了時刻は開始より後にしてください。"};
  if (label.length > 60 || memo.length > 200) return {error:"too_long", message:"内容は60文字、メモは200文字までです。"};
  return {kind:kind, name:kind === "業務" ? "" : name, start:start, end:end, from:from, to:to, label:kind === "業務" ? label : "", memo:memo};
}
function findShiftRow_(sheet, id) {
  const lastRow = sheet.getLastRow(); if (lastRow < 2 || !id) return 0;
  const found = sheet.getRange(2, 1, lastRow - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return found ? found.getRow() : 0;
}
function addShift_(ss, params) {
  const id = String(params.id || "");
  if (!/^[A-Za-z0-9-]{8,64}$/.test(id)) return {ok:false, action:"shiftAdd", error:"invalid_id", message:"IDが不正です。"};
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = getOrCreateShiftSheet_(ss);
    if (findShiftRow_(sheet, id)) return {ok:true, action:"shiftAdd", id:id, duplicate:true};
    const v = validateShift_(params, readRoster_(ss).users);
    if (v.error) return {ok:false, action:"shiftAdd", error:v.error, message:v.message};
    const by = String(params.by || "").trim().slice(0, 20) || "不明";
    const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([id, v.kind, v.name, v.start, v.end, v.from, v.to, v.label, v.memo, by, now, "", "", ""]);
    SpreadsheetApp.flush();
    clearShiftCache_(v.start, v.end);
    return {ok:true, action:"shiftAdd", id:id, entry:Object.assign({id:id, by:by, at:now, deleted:false}, v)};
  } finally { lock.releaseLock(); }
}
function deleteShift_(ss, params) {
  const id = String(params.id || ""), by = String(params.by || "").trim().slice(0, 20);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = ss.getSheetByName(SHIFT_SHEET_NAME);
    const row = sheet ? findShiftRow_(sheet, id) : 0;
    if (!row) return {ok:false, action:"shiftDelete", error:"not_found", message:"その予定は見つかりません。"};
    const e = shiftEntry_(sheet.getRange(row, 1, 1, SHIFT_HEADERS.length).getDisplayValues()[0], row);
    if (e.deleted) return {ok:true, action:"shiftDelete", id:id, duplicate:true};
    const admin = validateAdmin_(params, "shiftDelete") === null;
    // Anyone may remove what they entered or what carries their own name; the rest needs the admin.
    if (!admin && !(by && (e.name === by || e.by === by))) return {ok:false, action:"shiftDelete", error:"forbidden", message:"自分の予定以外は管理者だけが削除できます。"};
    const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(row, 12, 1, 3).setValues([["削除", by || "管理者", now]]);
    SpreadsheetApp.flush();
    clearShiftCache_(e.start, e.end);
    return {ok:true, action:"shiftDelete", id:id};
  } finally { lock.releaseLock(); }
}

// ===== お知らせ =====
function getOrCreateNoticeSheet_(ss) {
  let sheet = ss.getSheetByName(NOTICE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(NOTICE_SHEET_NAME);
    sheet.appendRow(NOTICE_HEADERS);
    sheet.getRange(1, 1, 1, NOTICE_HEADERS.length).setFontWeight("bold");
    sheet.getRange("A:H").setNumberFormat("@");
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function noticeAuth_(params) {
  if (validateAdmin_(params, "notice") === null) return true;
  const expected = PropertiesService.getScriptProperties().getProperty("NOTICE_TOKEN");
  return !!expected && expected.length >= 16 && String(params.noticeToken || "") === expected;
}
function readNotices_(ss) {
  // Notices ride along with every read; a problem here must never break attendance itself.
  try {
    const sheet = ss && typeof ss.getSheetByName === "function" ? ss.getSheetByName(NOTICE_SHEET_NAME) : null;
    if (!sheet || sheet.getLastRow() < 2) return [];
    const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, NOTICE_HEADERS.length).getDisplayValues()
      .filter(r => r[0] && String(r[6] || "") !== "削除")
      .map(r => ({id:String(r[0]), text:String(r[1] || ""), from:shiftDate_(r[2]), until:shiftDate_(r[3]), by:String(r[4] || ""), at:r[5] ? indexTime_(r[5]) : ""}))
      .filter(n => n.text && (!n.from || n.from <= today) && (!n.until || n.until >= today))
      .sort((a, b) => b.at.localeCompare(a.at)).slice(0, 10);
  } catch (e) { return []; }
}
function addNotice_(ss, params) {
  if (!noticeAuth_(params)) return {ok:false, action:"noticeAdd", error:"unauthorized", message:"お知らせの掲載には管理者認証が必要です。"};
  const text = String(params.text || "").replace(/\r/g, "").trim();
  if (!text) return {ok:false, action:"noticeAdd", error:"missing_text", message:"本文を入力してください。"};
  if (text.length > 500) return {ok:false, action:"noticeAdd", error:"too_long", message:"本文は500文字までです。"};
  const from = shiftDate_(params.from), until = shiftDate_(params.until);
  if (from && until && until < from) return {ok:false, action:"noticeAdd", error:"invalid_date", message:"掲載終了は掲載開始以降にしてください。"};
  const id = /^[A-Za-z0-9-]{8,64}$/.test(String(params.id || "")) ? String(params.id) : Utilities.getUuid();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = getOrCreateNoticeSheet_(ss);
    if (findShiftRow_(sheet, id)) return {ok:true, action:"noticeAdd", id:id, duplicate:true};
    const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([id, text, from, until, String(params.by || "管理者").trim().slice(0, 20) || "管理者", now, "", ""]);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove("attendance:recent");
    return {ok:true, action:"noticeAdd", id:id, notice:{id:id, text:text, from:from, until:until, by:String(params.by || "管理者").trim().slice(0, 20) || "管理者", at:now}};
  } finally { lock.releaseLock(); }
}
function deleteNotice_(ss, params) {
  if (!noticeAuth_(params)) return {ok:false, action:"noticeDelete", error:"unauthorized", message:"お知らせの削除には管理者認証が必要です。"};
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = ss.getSheetByName(NOTICE_SHEET_NAME);
    const row = sheet ? findShiftRow_(sheet, String(params.id || "")) : 0;
    if (!row) return {ok:false, action:"noticeDelete", error:"not_found", message:"そのお知らせは見つかりません。"};
    sheet.getRange(row, 7, 1, 2).setValues([["削除", Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss")]]);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove("attendance:recent");
    return {ok:true, action:"noticeDelete", id:String(params.id)};
  } finally { lock.releaseLock(); }
}
