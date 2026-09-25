const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),crypto=require('node:crypto');
// ---- Apps Script side: master sheet + roster sheets; the summary sheets are created by the code under test ----
function chain(obj){ return new Proxy(obj,{get(t,p){ if(p in t) return t[p]; if(p==='then') return undefined; return ()=>chain(t); }}); }
function makeSheet(name,rows){ rows=rows||[]; const cells=new Map(); let frozen=0;
  const sheet={rows,cells,getName:()=>name,getLastRow:()=>rows.length,appendRow:r=>rows.push(r.slice()),setFrozenRows(n){frozen=n;},getFrozenRows:()=>frozen,
    getRange(r,c,n=1,w=1){
      if(typeof r==='string'){ const key=r; return chain({getValue:()=>key==='I1'?(rows[0]?rows[0][8]:''):(cells.get(key)??''),setValue(v){cells.set(key,v);return this;},setFormula(f){cells.set(key,f);return this;},getFormula:()=>cells.get(key)??''}); }
      return chain({getValues:()=>rows.slice(r-1,r-1+n).map(row=>Array.from({length:w},(_,i)=>row[c-1+i]==null?'':row[c-1+i])),
        getDisplayValues:()=>rows.slice(r-1,r-1+n).map(row=>Array.from({length:w},(_,i)=>row[c-1+i]==null?'':String(row[c-1+i]))),
        setValues(vals){ while(rows.length<r-1+vals.length) rows.push([]); vals.forEach((vr,i)=>vr.forEach((v,j)=>{rows[r-1+i][c-1+j]=v;})); return this; },
        clearContent(){ for(let i=r-1;i<r-1+n && i<rows.length;i++) for(let j=c-1;j<c-1+w;j++) rows[i][j]=''; while(rows.length>1 && rows[rows.length-1].every(v=>v===''||v==null)) rows.pop(); return this; },
        createTextFinder(value){const found=rows.slice(r-1,r-1+n).flatMap((row,i)=>String(row[c-1])===String(value)?[{getRow:()=>r+i}]:[]);return {matchEntireCell(){return this},matchCase(){return this},findNext:()=>found[0]||null,findAll:()=>found};}
      });
    }};
  return chain(sheet); }
const ALLOC='[{"id":"hw_production","minutes":60},{"id":"hwa_text","minutes":105}]';
const master=makeSheet('勤怠マスタ',[['t','名前','区分','打刻日時','対象月','システムID','交通機関','備考','業務区分','実働','','','','','','','','','','','業務配分データ'],
  ['t','喜多','出勤','2026-09-01 9:00:00','2026-09','k1','バス・電車','','','','','','','','','','','','','',''],
  ['t','喜多','休憩開始','2026-09-01 12:00:00','2026-09','k2','','','','','','','','','','','','','','',''],
  ['t','喜多','休憩終了','2026-09-01 12:30:00','2026-09','k3','','','','','','','','','','','','','','',''],
  ['t','喜多','退勤','2026-09-01 14:00:00','2026-09','k4','','','','','','','','','','','','','','',''],
  ['t','喜多','出勤','2026-09-02 9:00:00','2026-09','k5','自転車','','','','','','','','','','','','','',''],
  ['t','喜多','退勤','2026-09-02 13:30:00','2026-09','k6','','','','','','','','','','','','','','',''],
  ['t','田中','出勤','2026-09-03 9:00:00','2026-09','t1','','','業務配分','','','','','','','','','','','',''],
  ['t','田中','退勤','2026-09-03 11:45:00','2026-09','t2','','','業務配分','','','','','','','','','','','',ALLOC]]);
const placeholder=makeSheet('月次集計ダッシュボード',[['後ほど一目で見られるように作成']]);
const sheets=[master,placeholder,makeSheet('喜多_勤務表',[['対象月','2026-09','喜多','','','','','往復交通費',636]]),makeSheet('田中_勤務表',[['対象月','2026-09','田中','','','','','往復交通費',0]])];
const ss={getSheetByName:n=>sheets.find(s=>s.getName()===n)||null,insertSheet:n=>{const s=makeSheet(n);sheets.push(s);return s;},getSheets:()=>sheets,setActiveSheet(){},moveActiveSheet(){}};
const cache=new Map(),props=new Map();
const pad=n=>String(n).padStart(2,'0');
const gas={CacheService:{getScriptCache:()=>({get:k=>cache.get(k)||null,put:(k,v)=>cache.set(k,v),remove:k=>cache.delete(k)})},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v),deleteProperty:k=>props.delete(k)})},
  Utilities:{getUuid:()=>crypto.randomUUID(),formatDate:(d,tz,f)=>f==='yyyy-MM'?`${d.getFullYear()}-${pad(d.getMonth()+1)}`:f==='yyyy-MM-dd'?`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`:`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`},
  LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},SpreadsheetApp:chain({flush(){},openById:()=>ss}),
  ContentService:{createTextOutput:s=>({setMimeType(){return {getContent:()=>s}}}),MimeType:{JSON:'json',JAVASCRIPT:'js'}},Date,JSON,Math,console};
vm.createContext(gas);vm.runInContext(fs.readFileSync('Code.gs','utf8'),gas);
// The stats mirror the app: 喜多 9h over two days (one paid transport day), 田中 165min from allocations.
const data=gas.readData_(ss,master,{scope:'month',month:'2026-09'});assert.equal(data.ok,true);assert.equal(data.logs.length,8);
const st=gas.monthlyStats_(data.logs,data.users,'2026-09');
assert.equal(st['喜多'].totalMinutes,540);assert.equal(st['喜多'].daysCount,2);assert.equal(st['喜多'].paidTransportDays,1);
assert.equal(st['田中'].totalMinutes,165);assert.equal(st['田中'].categoryMinutes['ホームワイン'],60);assert.equal(st['田中'].categoryMinutes['アカデミー'],105);
// Writing creates both sheets, sorts members by hours, appends a 合計 row, and formats via the data sheet.
assert.equal(gas.writeMonthlySummary_(ss,'2026-09'),3);
const dataSheet=ss.getSheetByName('月次集計_データ'),view=ss.getSheetByName('月次集計ダッシュボード');assert(dataSheet&&view);assert.equal(view,placeholder); // the hand-made placeholder is configured in place
assert.equal(JSON.stringify(dataSheet.rows[0]),JSON.stringify(['月','名前','出勤日数','勤務時間','時間(小数)','目標80H比','交通費支給日','往復交通費','交通費合計','アカデミー','ホームワイン','その他（WT業務）','更新日時']));
const kita=dataSheet.rows[1],tanaka=dataSheet.rows[2],total=dataSheet.rows[3];
assert.equal(JSON.stringify(kita.slice(0,12)),JSON.stringify(['2026-09','喜多',2,540/1440,9,540/4800,1,636,636,0,0,0]));
assert.equal(JSON.stringify(tanaka.slice(0,12)),JSON.stringify(['2026-09','田中',1,165/1440,2.75,165/4800,0,0,0,105/1440,60/1440,0]));
assert.equal(JSON.stringify(total.slice(0,12)),JSON.stringify(['2026-09','合計',3,705/1440,11.75,'',1,'',636,105/1440,60/1440,0]));
assert.match(String(kita[12]),/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/);assert.equal(dataSheet.rows.length,4);
assert.equal(view.getRange('B1').getValue(),'2026-09');assert.match(view.getRange('A8').getFormula(),/FILTER\(\{'月次集計_データ'!B2:B/);assert.match(view.getRange('E5').getFormula(),/TEXTJOIN/);assert.equal(view.getFrozenRows(),7);
// An empty month gets zero rows for everyone, sorted below the newer month; rewriting a month never duplicates it.
assert.equal(gas.writeMonthlySummary_(ss,'2026-08'),3);
assert.equal(dataSheet.rows.length,7);assert.equal(dataSheet.rows[1][0],'2026-09');assert.equal(dataSheet.rows[4][0],'2026-08');assert.equal(dataSheet.rows[4][3],0);assert.equal(dataSheet.rows[6][1],'合計');
assert.equal(gas.writeMonthlySummary_(ss,'2026-09'),3);assert.equal(dataSheet.rows.length,7);assert.equal(dataSheet.rows[1][1],'喜多');
assert.equal(view.getRange('B1').getValue(),'2026-09'); // the chosen month is left alone once set
// Months to refresh: a punch refreshes its own month and the current one; a deletion also refreshes the previous month.
{ const m=gas.summaryMonthsFor_('add',{month:'2026-07',time:'2026-07-01 9:00:00'});assert.equal(m.length,2);assert.equal(m[1],'2026-07');
  const d=gas.summaryMonthsFor_('deleteRecent',{});assert.equal(d.length,2); }
// Failed months are remembered and retried; the receipt-protected router refreshes after a punch is stored.
props.set('summary:dirty','2026-06');
let r=gas.refreshMonthlySummary_(['2026-09']);assert.equal(JSON.stringify(r.refreshed.sort()),JSON.stringify(['2026-06','2026-09']));assert.equal(props.get('summary:dirty'),'');
assert.equal(dataSheet.rows.length,10);
{ const receipt=crypto.randomUUID();
  const res=JSON.parse(gas.handleRequest_({parameter:{action:'add',receipt,requestId:'req-1',name:'喜多',type:'出勤',time:'2026-09-04 9:00:00',month:'2026-09',transport:'バス・電車',allocations:'[]'}},true).getContent());
  assert.equal(res.ok,true);
  const row=dataSheet.rows.find(x=>x[0]==='2026-09'&&x[1]==='喜多');assert.equal(row[2],3);assert.equal(row[6],2);assert.equal(row[8],1272); }
// The rebuild action needs the notice token and accepts several months at once.
{ const receipt=crypto.randomUUID();
  assert.equal(JSON.parse(gas.handleRequest_({parameter:{action:'summaryRebuild',receipt,month:'2026-09'}},true).getContent()).error,'unauthorized');
  props.set('NOTICE_TOKEN','abcdefghijklmnop123456');
  assert.equal(gas.rebuildSummary_(ss,{month:'2026-13',noticeToken:'abcdefghijklmnop123456'}).error,'invalid_month');
  const res=JSON.parse(gas.handleRequest_({parameter:{action:'summaryRebuild',receipt:crypto.randomUUID(),month:'2026-09,2026-05',noticeToken:'abcdefghijklmnop123456'}},true).getContent());
  assert.equal(res.ok,true);assert.equal(res.months['2026-05'],3);assert.equal(dataSheet.rows.length,13);assert.equal(dataSheet.rows[dataSheet.rows.length-1][0],'2026-05'); }
console.log('test-summary: OK');
