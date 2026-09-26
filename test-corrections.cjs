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
        setValues(v){v.forEach((vr,i)=>vr.forEach((x,j)=>{rows[r-1+i][c-1+j]=x;}));return this;},
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
const post=p=>JSON.parse(gas.handleRequest_({parameter:Object.assign({action:'add',receipt:crypto.randomUUID(),month:p.time.slice(0,7),correction:'1'},p)},true).getContent());
const logs=()=>master.rows.slice(1).map((r,i)=>gas.logFromRow_(r.map(v=>v==null?'':String(v)),i+2));
// ---- The fit check: a forgotten punch goes in only where it cancels nothing already recorded ----
{ const L=[{id:'a',name:'X',type:'出勤',time:at(5,9)},{id:'b',name:'X',type:'退勤',time:at(5,15)},{id:'c',name:'X',type:'出勤',time:at(6,9)},{id:'d',name:'X',type:'休憩開始',time:at(6,12)}];
  const fit=(type,time,cat)=>gas.correctionFit_(L,'X',type,time,cat);
  assert.equal(fit('退勤',at(6,16)),null);                                   // forgot to clock out after a break
  assert.equal(fit('休憩終了',at(6,12,30)),null);
  assert.match(fit('退勤',at(6,11)).message,/12:00 の休憩開始が無効/);          // would orphan the later break
  assert.match(fit('出勤',at(5,10)).message,/すでに勤務中/);
  assert.match(fit('退勤',at(7,10)).message,/出勤の記録がありません/);
  assert.match(fit('退勤',at(5,9)).message.replace(/\s/g,''),/./);             // 退勤 at the very clock-in time of a closed shift is refused
  assert.match(fit('退勤',jst(new Date(at(6,9).replace(' ','T')+'+09:00').getTime()+19*H)).message,/18時間を超える/);
  assert.match(fit('休憩終了',at(5,12)).message,/休憩中ではありません/);
  assert.match(fit('休憩開始',at(6,12,10)).message,/すでに休憩中/);
  assert.match(fit('退勤',at(5,15)).message,/同じ時刻の退勤/);
  assert.equal(fit('休憩開始',at(5,12)),null);                               // a forgotten break inside a finished shift is fine by punches…
  assert.match(fit('休憩開始',at(5,12),'業務配分').message,/退勤済み/);            // …but not after 業務配分 was entered at checkout
}
// ---- A correction is saved at once, marked in 修正, memo and red, and refuses what cannot be ---
const before=master.rows.length;
assert.equal(post({requestId:'x1',name:'橋本',type:'退勤',time:jst(NOW+H)}).message,'これからの時刻には修正依頼できません。');
assert.match(post({requestId:'x2',name:'橋本',type:'退勤',time:at(40,15)}).message,/31日より前/);
assert.match(post({requestId:'x3',name:'橋本',type:'退勤',time:at(3,15)}).message,/出勤の記録がありません/);
assert.match(post({requestId:'x4',name:'橋本',type:'早退',time:at(2,15)}).message,/出勤・退勤・休憩開始・休憩終了/);
assert.equal(master.rows.length,before,'refused corrections write nothing');
let r=post({requestId:'fix-1',name:'橋本',type:'退勤',time:at(2,15,30),reason:'退勤を押し忘れました\n｜すみません'});
assert.equal(r.ok,true);assert.match(r.corrected,/^修正 \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/);
assert.match(r.memo,/^【修正】退勤を押し忘れました すみません（\d+\/\d+ \d+:\d{2} 入力）$/);
const saved=master.rows[master.rows.length-1];assert.equal(saved[5],'fix-1');assert.equal(saved[21],r.corrected);assert.equal(saved[7],r.memo);
assert(fontColors.some(f=>f.sheet==='勤怠マスタ'&&f.row===master.rows.length&&f.cols===22&&f.color==='#d32f2f'),'the corrected row is red');
assert.equal(post({requestId:'fix-1',name:'橋本',type:'退勤',time:at(2,15,30)}).duplicate,true); // a resend is not a second row
{ const n=master.rows.length,again=post({requestId:'fix-1b',name:'橋本',type:'退勤',time:at(2,15,30)});assert(again.duplicate===true||/同じ時刻の退勤/.test(again.message));assert.equal(master.rows.length,n,'the same correction twice is one row'); }
// The member sheet gets its red rules once: each punch column by its kind, hours/break/備考 for any correction.
{ const hs=ss.getSheetByName('橋本_勤務表').rules;assert.equal(hs.length,5);
  assert.equal(hs[1].f,'=REGEXMATCH($J4,"退勤: 【修正】")');assert.equal(hs[1].ranges[0].col,3);assert.equal(hs[1].ranges[0].row,4);assert.equal(hs[1].ranges[0].rows,40);
  assert.equal(JSON.stringify(hs[4].ranges.map(x=>x.col)),'[4,7,10]');assert.equal(hs[0].color,'#d32f2f');
  assert.equal(gas.ensureCorrectionFormats_(ss,'橋本'),false);assert.equal(ss.getSheetByName('橋本_勤務表').rules.length,5); }
// A forgotten break in a finished shift (not 業務配分) goes in as a pair.
assert.equal(post({requestId:'fix-2',name:'橋本',type:'休憩開始',time:at(4,12)}).ok,true);
assert.equal(post({requestId:'fix-3',name:'橋本',type:'休憩終了',time:at(4,12,30)}).ok,true);
// 業務配分: the checkout needs the split of the hours up to the corrected time, taken from the sheet (the old open shift).
{ const wrong=post({requestId:'fix-4',name:'田中',type:'退勤',time:at(2,13),category:'業務配分',allocations:JSON.stringify(ALLOC(0,60))});
  assert.equal(wrong.error,'allocation_mismatch');assert.equal(wrong.actualMinutes,180);
  const ok=post({requestId:'fix-5',name:'田中',type:'退勤',time:at(2,13),category:'業務配分',allocations:JSON.stringify(ALLOC(1,180))});
  assert.equal(ok.ok,true);assert.match(ok.memo,/^【修正】押し忘れ（.*入力）｜HWのお問い合わせ（180分）$/);
  assert.match(post({requestId:'fix-6',name:'田中',type:'休憩開始',time:at(3,10),category:'業務配分'}).message,/退勤済み/); }
// ---- Reads: the flag rides along, and the punch screen sees a backdated correction ----
{ for(let k=1;k<=60;k++) master.rows.push(row('喜多',k%2?'出勤':'退勤',jst(NOW-k*5*60000),'k'+k));
  gas.invalidateAttendanceIndex_();
  const recent=gas.readData_(ss,master,{scope:'recent',fresh:'1'});
  const fix=recent.logs.find(l=>l.id==='fix-1');assert(fix,'the correction is in the punch-screen read');assert.match(fix.corrected,/^修正 /);
  const month=gas.readData_(ss,master,{scope:'month',month:at(2,15,30).slice(0,7)});assert(month.logs.some(l=>l.id==='fix-1'&&l.corrected)); }
// ---- Summary: corrections are counted per member and in total ----
{ const L=logs(),m=at(2,15,30).slice(0,7);const st=gas.monthlyStats_(L,['橋本','田中','喜多'],m);
  const expectH=L.filter(l=>l.name==='橋本'&&l.month===m&&l.corrected).length;assert.equal(st['橋本'].corrections,expectH);assert(expectH>=1);assert.equal(st['喜多'].corrections,0); }
// ---- Maintenance: pre-applying the rules to every member sheet needs the admin token ----
assert.equal(gas.setupCorrectionFormats_(ss,{}).error,'unauthorized');
props.set('NOTICE_TOKEN','abcdefghijklmnop123456');
{ const res=gas.setupCorrectionFormats_(ss,{noticeToken:'abcdefghijklmnop123456'});assert.equal(res.ok,true);assert.equal(JSON.stringify(res.sheets),JSON.stringify(['喜多_勤務表'])); }
console.log('test-corrections: OK');
