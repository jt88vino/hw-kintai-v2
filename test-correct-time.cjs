const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),crypto=require('node:crypto');
// ---- Times anchored to today's JST midnight so the scenario has fixed clock times whenever it runs ----
const H=3600000,NOW=Date.now(),pad=n=>String(n).padStart(2,'0');
function jst(ms){const d=new Date(ms+9*H);return `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;}
const midnight=(()=>{const d=new Date(NOW+9*H);return Date.UTC(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())-9*H;})();
const at=(daysAgo,h,m=0)=>jst(midnight-daysAgo*24*H+h*H+m*60000);
// ---- Spreadsheet mock: master + member sheets with conditional-format storage ----
const fontColors=[];
function makeSheet(name,rows){ rows=rows||[]; let rules=[];
  const sheet={rows,getName:()=>name,getLastRow:()=>rows.length,appendRow:r=>rows.push(r.slice()),setFrozenRows(){},getParent:()=>ss,
    getConditionalFormatRules:()=>rules.slice(),setConditionalFormatRules(r){rules=r.slice();},get rules(){return rules;},
    getRange(r,c,n=1,w=1){
      if(typeof r==='string') return {getValue:()=>r==='I1'?(rows[0]?rows[0][8]:''):'',setNumberFormat(){return this},setFontWeight(){return this}};
      return {row:r,col:c,rows:n,cols:w,
        getDisplayValues:()=>rows.slice(r-1,r-1+n).map(row=>Array.from({length:w},(_,i)=>row[c-1+i]==null?'':String(row[c-1+i]))),
        getValues:()=>rows.slice(r-1,r-1+n).map(row=>Array.from({length:w},(_,i)=>row[c-1+i]==null?'':row[c-1+i])),
        setValue(x){rows[r-1][c-1]=x;return this;},setValues(v){v.forEach((vr,i)=>vr.forEach((x,j)=>{rows[r-1+i][c-1+j]=x;}));return this;},
        setFontColor(color){fontColors.push({sheet:name,row:r,cols:w,color});return this;},setFormula(){return this},setNumberFormat(){return this},setFontWeight(){return this},
        createTextFinder(value){const found=rows.slice(r-1,r-1+n).flatMap((row,i)=>String(row[c-1])===String(value)?[{getRow:()=>r+i}]:[]);return {matchEntireCell(){return this},matchCase(){return this},findNext:()=>found[0]||null,findAll:()=>found};}};
    }};
  return sheet; }
const row=(name,type,time,id,category='',extra={})=>{const r=Array(22).fill('');Object.assign(r,{0:'t',1:name,2:type,3:time,4:time.slice(0,7),5:id,8:category});Object.entries(extra).forEach(([k,v])=>r[k]=v);return r;};
const ALLOC=(first,minutes)=>{const items=JSON.parse(fs.readFileSync('Code.gs','utf8').match(/const ALLOCATION_ITEMS = (\[.*?\]);/)[1]);return items.map((i,k)=>({id:i.id,minutes:k===first?minutes:0,memo:''}));};
const master=makeSheet('勤怠マスタ',[['タイムスタンプ','名前','区分','打刻日時','対象月','システムID','交通機関','備考','業務区分','実働時間'],
  row('橋本','出勤',at(2,8,30),'h1'),                                        // 2 days ago: clocked in, forgot to clock out
  row('橋本','出勤',at(1,9),'h2'),row('橋本','退勤',at(1,15),'h3'),            // yesterday: complete
  row('橋本','出勤',at(4,9),'h4'),row('橋本','退勤',at(4,15),'h5'),            // 4 days ago: complete, break forgotten
  row('田中','出勤',at(3,9),'t1','業務配分'),row('田中','退勤',at(3,12),'t2','業務配分',{20:JSON.stringify(ALLOC(0,180))}),
  row('田中','出勤',at(2,10),'t3','業務配分')]);                              // 2 days ago: forgot to clock out
const memberHeads=['日付','出勤','退勤','休憩時間','休憩開始','休憩終了','実働時間','発生交通費','交通機関','備考'];
const member=n=>makeSheet(n+'_勤務表',[['対象月','2026-09',n,'','','','','往復交通費',0],[],memberHeads]);
const sheets=[master,member('橋本'),member('田中'),member('喜多')];
const ss={getSheetByName:n=>sheets.find(s=>s.getName()===n)||null,insertSheet:n=>{const s=makeSheet(n);sheets.push(s);return s;},getSheets:()=>sheets,setActiveSheet(){},moveActiveSheet(){}};
const cache=new Map(),props=new Map();
const rule=()=>{const b={whenFormulaSatisfied(f){b.f=f;return b},setFontColor(c){b.color=c;return b},setBold(){return b},setBackground(){return b},setRanges(r){b.ranges=r;return b},
  build(){const f=b.f,ranges=b.ranges,color=b.color;return {f,ranges,color,getBooleanCondition:()=>({getCriteriaValues:()=>[f]})};}};return b;};
const gas={CacheService:{getScriptCache:()=>({get:k=>cache.get(k)||null,put:(k,v)=>cache.set(k,v),remove:k=>cache.delete(k)})},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v),deleteProperty:k=>props.delete(k)})},
  Utilities:{getUuid:()=>crypto.randomUUID(),formatDate:(d,tz,f)=>jst(d.getTime()).slice(0,f==='yyyy-MM'?7:f==='yyyy-MM-dd'?10:16)},
  LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},SpreadsheetApp:{flush(){},openById:()=>ss,newConditionalFormatRule:rule},
  ContentService:{createTextOutput:s=>({setMimeType(){return {getContent:()=>s}}}),MimeType:{JSON:'json',JAVASCRIPT:'js'}},Date,JSON,Math,console};
vm.createContext(gas);vm.runInContext(fs.readFileSync('Code.gs','utf8'),gas);

// ---- 時刻の修正（新しい社内アプリで承認された「押した時刻を直す」）----
const fix=p=>JSON.parse(gas.handleRequest_({parameter:Object.assign({action:'correctTime',receipt:crypto.randomUUID()},p)},true).getContent());
const rowOf=id=>master.rows.find(r=>r[5]===id);
master.rows.push(row('喜多','出勤',at(0,9,26),'k1'));
const width=gas.MASTER_HEADERS?gas.MASTER_HEADERS.length:22;
// GET では受け付けない
assert.equal(JSON.parse(gas.handleRequest_({parameter:{action:'correctTime',id:'k1',name:'喜多',type:'出勤',time:at(0,9)}},false).getContent()).ok,false);
assert.equal(rowOf('k1')[3],at(0,9,26));
// 名前・区分が合わない、見つからない、未来、12時間以上は受け付けない
assert.equal(fix({id:'k1',name:'橋本',type:'出勤',time:at(0,9)}).error,'mismatch');
assert.equal(fix({id:'k1',name:'喜多',type:'退勤',time:at(0,9)}).error,'mismatch');
assert.equal(fix({id:'nope',name:'喜多',type:'出勤',time:at(0,9)}).error,'not_found');
assert.equal(fix({id:'k1',name:'喜多',type:'出勤',time:jst(NOW+H)}).error,'invalid_time');
assert.equal(fix({id:'k1',name:'喜多',type:'出勤',time:at(1,20)}).error,'too_far');
assert.equal(rowOf('k1')[3],at(0,9,26),'refused fixes change nothing');
// 直すと、時刻・備考の【修正】（元→新）・修正列・赤字が入る
let f=fix({id:'k1',name:'喜多',type:'出勤',time:at(0,9),reason:'9:00に来ていた\n｜すみません'});
assert.equal(f.ok,true);assert.equal(f.action,'correctTime');
const k=rowOf('k1');
assert.equal(k[3],at(0,9));
assert.match(k[7],/^【修正】9:00に来ていた すみません（09:26→09:00、\d+\/\d+ \d+:\d{2} 入力）$/);
assert.match(k[width-1],/^修正 \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/);
assert.ok(fontColors.some(c=>c.sheet==='勤怠マスタ'&&c.row===master.rows.indexOf(k)+1&&c.color==='#d32f2f'));
// 同じ内容をもう一度送っても1回分（備考は増えない）
const memo=k[7];
assert.equal(fix({id:'k1',name:'喜多',type:'出勤',time:at(0,9)}).duplicate,true);
assert.equal(rowOf('k1')[7],memo);
// 結果は受付番号でも読める（新しいアプリが返事を読めなかったとき）
const receipt=crypto.randomUUID();
gas.handleRequest_({parameter:{action:'correctTime',receipt,id:'k1',name:'喜多',type:'出勤',time:at(0,9,5)}},true);
assert.equal(JSON.parse(gas.handleRequest_({parameter:{action:'adminResult',receipt}},false).getContent()).ok,true);
assert.equal(rowOf('k1')[3],at(0,9,5));
console.log('correct-time ok');
