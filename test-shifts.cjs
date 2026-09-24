const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),crypto=require('node:crypto');
// ---- Apps Script side: a two-sheet spreadsheet mock (roster sheets + the シフト sheet the code creates) ----
function makeSheet(name,rows){ rows=rows||[]; return {rows,getName:()=>name,getLastRow:()=>rows.length,appendRow:r=>rows.push(r.slice()),setFrozenRows(){},
  getRange(r,c,n=1,w=1){
    if(typeof r==='string') return r==='I1'?{getValue:()=>rows[0]?rows[0][8]:''}:{setNumberFormat(){return this},setFontWeight(){return this}};
    return {getDisplayValues:()=>rows.slice(r-1,r-1+n).map(row=>Array.from({length:w},(_,i)=>row[c-1+i]==null?'':String(row[c-1+i]))),
      setValues(vals){vals.forEach((vr,i)=>vr.forEach((v,j)=>{rows[r-1+i][c-1+j]=v;}));return this;},setFontWeight(){return this},setNumberFormat(){return this},
      createTextFinder(value){const found=rows.slice(r-1,r-1+n).flatMap((row,i)=>String(row[c-1])===String(value)?[{getRow:()=>r+i}]:[]);return {matchEntireCell(){return this},matchCase(){return this},findNext:()=>found[0]||null,findAll:()=>found};}};
  }};}
const sheets=[makeSheet('勤怠マスタ',[['t']]),makeSheet('喜多_勤務表',[['対象月','2026-09','喜多','','','','','往復交通費',636]]),makeSheet('橋本_勤務表',[['対象月','2026-09','橋本','','','','','往復交通費',0]]),makeSheet('松井_勤務表',[['対象月','2026-09','松井','','','','','往復交通費',0]]),makeSheet('田中_勤務表',[['対象月','2026-09','田中','','','','','往復交通費',0]]),makeSheet('鈴木_勤務表',[['対象月','2026-09','鈴木','','','','','往復交通費',0]]),makeSheet('長谷川_勤務表',[['対象月','2026-09','長谷川','','','','','往復交通費',0]])];
const ss={getSheetByName:n=>sheets.find(s=>s.getName()===n)||null,insertSheet:n=>{const s=makeSheet(n);sheets.push(s);return s;},getSheets:()=>sheets};
const cache=new Map(),props=new Map();
const pad=n=>String(n).padStart(2,'0');
const gas={CacheService:{getScriptCache:()=>({get:k=>cache.get(k)||null,put:(k,v)=>cache.set(k,v),remove:k=>cache.delete(k)})},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v),deleteProperty:k=>props.delete(k)})},
  Utilities:{getUuid:()=>crypto.randomUUID(),formatDate:(d,tz,f)=>f==='yyyy-MM-dd'?`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`:f==='HH:mm'?`${pad(d.getHours())}:${pad(d.getMinutes())}`:`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`},
  LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},SpreadsheetApp:{flush(){},openById:()=>ss},
  ContentService:{createTextOutput:s=>({setMimeType(){return {getContent:()=>s}}}),MimeType:{JSON:'json',JAVASCRIPT:'js'}}};
vm.createContext(gas);vm.runInContext(fs.readFileSync('Code.gs','utf8'),gas);
const add=p=>gas.addShift_(ss,Object.assign({id:crypto.randomUUID()},p));
let r=add({kind:'シフト',name:'橋本',start:'2026-10-03',from:'9:30',to:'15:00',by:'橋本'});assert.equal(r.ok,true);assert.equal(r.entry.from,'09:30');assert.equal(r.entry.end,'2026-10-03');
const sh=ss.getSheetByName('シフト');assert.equal(sh.rows.length,2);assert.equal(sh.rows[0][0],'ID');assert.equal(sh.rows[1][9],'橋本');
const dupId=crypto.randomUUID();assert.equal(gas.addShift_(ss,{id:dupId,kind:'休み',name:'鈴木',start:'2026-10-05',by:'鈴木'}).ok,true);assert.equal(gas.addShift_(ss,{id:dupId,kind:'休み',name:'鈴木',start:'2026-10-05',by:'鈴木'}).duplicate,true);assert.equal(sh.rows.length,3);
assert.equal(add({kind:'なにか',name:'橋本',start:'2026-10-03'}).error,'invalid_kind');
assert.equal(add({kind:'シフト',start:'2026-10-03',from:'09:00',to:'14:00'}).error,'missing_name');
assert.equal(add({kind:'シフト',name:'誰か',start:'2026-10-03',from:'09:00',to:'14:00'}).error,'unknown_name');
assert.equal(add({kind:'シフト',name:'橋本',start:'2026-10-03'}).error,'missing_time');
assert.equal(add({kind:'シフト',name:'橋本',start:'2026-10-03',from:'15:00',to:'09:00'}).error,'invalid_time');
assert.equal(add({kind:'シフト',name:'橋本',start:'2026-10-03',end:'2026-10-04',from:'09:00',to:'14:00'}).error,'invalid_date');
assert.equal(add({kind:'業務',start:'2026-10-03'}).error,'missing_label');
assert.equal(add({kind:'業務',label:'x'.repeat(61),start:'2026-10-03'}).error,'too_long');
assert.equal(gas.addShift_(ss,{id:'x',kind:'休み',name:'鈴木',start:'2026-10-05'}).error,'invalid_id');
r=add({kind:'業務',label:'E・F瓶詰め',start:'2026-10-30',end:'2026-11-02',by:'長谷川'});assert.equal(r.ok,true);assert.equal(r.entry.name,'');
// Sheets may render what the app wrote as 2026/10/07 and 10:00:00; reads normalize it.
sh.rows.push([crypto.randomUUID(),'シフト','松井','2026/10/07','2026/10/07','10:00:00','15:00:00','','','松井','2026/10/01 9:00:00','','','']);
let m=gas.readShifts_(ss,{month:'2026-10',fresh:'1'});assert.equal(m.ok,true);assert.equal(m.entries.length,4);
const matsui=m.entries.find(e=>e.name==='松井');assert.equal(matsui.start,'2026-10-07');assert.equal(matsui.from,'10:00');assert.equal(matsui.at,'2026-10-01 09:00:00');
assert.equal(gas.readShifts_(ss,{month:'2026-11',fresh:'1'}).entries.length,1);assert.equal(gas.readShifts_(ss,{month:'2026-13'}).error,'invalid_month');
assert.equal(gas.readShifts_(ss,{month:'2026-10'}).entries.length,4);add({kind:'出勤',name:'田中',start:'2026-10-09',by:'田中'});assert.equal(gas.readShifts_(ss,{month:'2026-10'}).entries.length,5); // a write clears the cache
const hashimoto=m.entries.find(e=>e.name==='橋本');
assert.equal(gas.deleteShift_(ss,{id:hashimoto.id,by:'松井'}).error,'forbidden');
assert.equal(gas.deleteShift_(ss,{id:hashimoto.id,by:'橋本'}).ok,true);assert.equal(gas.deleteShift_(ss,{id:hashimoto.id,by:'橋本'}).duplicate,true);
const token='a'.repeat(72);cache.set('admin:'+token,'valid');assert.equal(gas.deleteShift_(ss,{id:matsui.id,by:'',adminToken:token}).ok,true);
assert.equal(gas.deleteShift_(ss,{id:'nope',by:'橋本'}).error,'not_found');
const after=gas.readShifts_(ss,{month:'2026-10',fresh:'1'});assert.equal(after.entries.length,3);assert(after.changes.some(c=>c.id===hashimoto.id&&c.deleted&&c.deletedBy==='橋本'));assert(after.changes.some(c=>c.id===matsui.id&&c.deleted&&c.deletedBy==='管理者'));
// Router: writes are POST with a receipt and are polled back; the month read is a plain GET.
const receipt=crypto.randomUUID(),id=crypto.randomUUID();
assert.equal(JSON.parse(gas.handleRequest_({parameter:{action:'shiftAdd',id,kind:'シフト',name:'鈴木',start:'2026-10-13',from:'09:00',to:'14:00',by:'鈴木'}},false).getContent()).error,'post_required');
assert.equal(JSON.parse(gas.handleRequest_({parameter:{action:'shiftAdd',receipt,id,kind:'シフト',name:'鈴木',start:'2026-10-13',from:'09:00',to:'14:00',by:'鈴木'}},true).getContent()).ok,true);
assert.equal(JSON.parse(gas.handleRequest_({parameter:{action:'adminResult',receipt}},false).getContent()).id,id);
assert.equal(JSON.parse(gas.handleRequest_({parameter:{action:'shift',month:'2026-10'}},false).getContent()).entries.length,4);
// ---- お知らせ: 管理者認証か NOTICE_TOKEN が要る。期間外・削除済みは配信されない ----
assert.equal(gas.addNotice_(ss,{text:'こんにちは'}).error,'unauthorized');
props.set('NOTICE_TOKEN','abcdefghijklmnop123456');
assert.equal(gas.addNotice_(ss,{text:'',noticeToken:'abcdefghijklmnop123456'}).error,'missing_text');
assert.equal(gas.addNotice_(ss,{text:'x',noticeToken:'wrong-token-wrong-token'}).error,'unauthorized');
let nt=gas.addNotice_(ss,{text:'10月からシフトはこのアプリで管理します。\n入力は本人で。',noticeToken:'abcdefghijklmnop123456'});assert.equal(nt.ok,true);assert.equal(nt.notice.by,'管理者');
assert.equal(gas.addNotice_(ss,{id:nt.id,text:'dup',noticeToken:'abcdefghijklmnop123456'}).duplicate,true);
const old=gas.addNotice_(ss,{text:'古い',until:'2020-01-01',adminToken:token,by:'管理者'});assert.equal(old.ok,true);
const future=gas.addNotice_(ss,{text:'未来',from:'2099-01-01',adminToken:token});assert.equal(future.ok,true);
assert.equal(gas.addNotice_(ss,{text:'逆',from:'2026-10-02',until:'2026-10-01',adminToken:token}).error,'invalid_date');
let notices=gas.readNotices_(ss);assert.equal(notices.length,1);assert.equal(notices[0].id,nt.id);assert.match(notices[0].text,/\n入力は本人で/);
assert.equal(gas.deleteNotice_(ss,{id:nt.id}).error,'unauthorized');assert.equal(gas.deleteNotice_(ss,{id:'zzz',adminToken:token}).error,'not_found');
assert.equal(gas.deleteNotice_(ss,{id:nt.id,noticeToken:'abcdefghijklmnop123456'}).ok,true);assert.equal(gas.readNotices_(ss).length,0);
cache.set('attendance:recent','{"stale":true}');gas.addNotice_(ss,{text:'再掲',noticeToken:'abcdefghijklmnop123456'});assert.equal(cache.has('attendance:recent'),false); // a new notice invalidates the cached read
// ---- Client side ----
const html=fs.readFileSync('index.html','utf8'),script=html.match(/<script>([\s\S]*?)<\/script>/)[1];const nodes=new Map();
function el(id){if(!nodes.has(id))nodes.set(id,{value:'',textContent:'',innerHTML:'',disabled:false,querySelectorAll:()=>[],classList:{add(){},remove(){},toggle(){}}});return nodes.get(id);}
const client={Date,URLSearchParams,Set,Map,console,localStorage:{getItem(){return null},setItem(){}},window:{addEventListener(){},crypto},document:{getElementById:el,querySelectorAll:()=>[]},setTimeout(){return 1},clearTimeout(){}};vm.createContext(client);
vm.runInContext(script.replace("if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();",`window.test={state,monthGridDays,chipText,sortEntries,nextShiftFor,todayShiftFor,shiftEntriesFor,personColor,JP_HOLIDAYS,queueShiftOp,runShiftOps,reconcileShiftOps,applyGasReadData,renderNotices,renderTodayShifts,upcomingEvents,renderEvents,setSend:fn=>{adminRequest=fn;},prepare:()=>{renderAll=()=>{};renderPunch=()=>{};renderShift=()=>{};showStatus=()=>{};getJSTDateTime=()=>({full:'2026-09-24 10:00:00',monthOnly:'2026-09',display:'10:00'});}};`),client);
const api=client.window.test;api.prepare();
api.applyGasReadData({ok:true,logs:[],notices:[{id:'n1',text:'掲示',by:'管理者',at:'2026-09-24 22:00:00'}]});assert.equal(api.state.notices.length,1);api.renderNotices();assert.match(el('noticeList').innerHTML,/掲示/);assert.match(el('noticeAdminList').innerHTML,/削除/);
let days=api.monthGridDays('2026-10');assert.equal(days[0],'2026-09-28');assert.equal(days[days.length-1],'2026-11-01');assert.equal(days.length,35);
days=api.monthGridDays('2026-09');assert.equal(days[0],'2026-08-31');assert.equal(days.length,35);assert.equal(api.monthGridDays('2026-11').length,42);
assert.equal(api.JP_HOLIDAYS['2026-09-21'],'敬老の日');assert.equal(api.JP_HOLIDAYS['2027-03-22'],'振替休日');
assert.equal(api.chipText({kind:'休み',name:'鈴木'}),'鈴木休み');assert.equal(api.chipText({kind:'有給',name:'牛嶋'}),'有給 牛嶋');assert.equal(api.chipText({kind:'シフト',name:'橋本',from:'09:30',to:'15:00'},true),'09:30-15:00 橋本');assert.equal(api.chipText({kind:'業務',label:'搬入'}),'📌 搬入');
assert.deepEqual(api.personColor('橋本'),api.personColor('橋本'));
const e=(id,kind,name,start,from,to,extra)=>Object.assign({id,kind,name,start,end:start,from:from||'',to:to||'',label:'',memo:'',by:name,at:''},extra||{});
api.state.shiftData['2026-09']={entries:[e('a','シフト','橋本','2026-09-24','09:30','15:00'),e('b','シフト','橋本','2026-09-30','10:00','15:00'),e('c','業務','','2026-09-28',null,null,{end:'2026-09-30',label:'小瓶搬入',by:'長谷川'})],changes:[]};
api.state.shiftData['2026-10']={entries:[e('d','シフト','橋本','2026-10-02','09:00','14:00')],changes:[]};
assert.equal(JSON.stringify(api.upcomingEvents(7).map(x=>x.id)),'["c"]');api.renderEvents();assert.match(el('eventList').innerHTML,/小瓶搬入/);assert.match(el('eventList').innerHTML,/9\/28\(月\)〜9\/30\(水\)/);
assert.equal(api.todayShiftFor('橋本').id,'a');api.state.users=['橋本'];api.renderTodayShifts();assert.match(el('todayShiftList').innerHTML,/09:30-15:00/);assert.match(el('todayShiftList').innerHTML,/未出勤/);assert.equal(api.todayShiftFor('松井'),null);
assert.equal(api.nextShiftFor('橋本','2026-09-25').id,'b');assert.equal(api.nextShiftFor('橋本','2026-10-01').id,'d');
assert.deepEqual(api.sortEntries(api.shiftEntriesFor('2026-09').filter(x=>x.start<='2026-09-30'&&x.end>='2026-09-30')).map(x=>x.id),['c','b']);
(async()=>{
 // An add goes out in the background with the entry fields and lands in the month once confirmed.
 let sent=[];api.setSend(async p=>{sent.push(p);return {ok:true,action:p.action,id:p.id};});
 api.queueShiftOp({op:'add',id:'n1',entry:e('n1','シフト','松井','2026-09-26','09:00','14:00'),by:'松井',sent:false});
 assert(api.shiftEntriesFor('2026-09').some(x=>x.id==='n1'&&x.pending));await api.runShiftOps();
 assert.equal(sent[0].action,'shiftAdd');assert.equal(sent[0].name,'松井');assert.equal(sent[0].from,'09:00');assert.equal(api.state.shiftOps.length,0);assert(api.state.shiftData['2026-09'].entries.some(x=>x.id==='n1'&&!x.pending));
 // A delete hides the entry at once and is dropped with a message when the server refuses it.
 api.setSend(async()=>{throw Object.assign(new Error('自分の予定以外は管理者だけが削除できます。'),{code:'forbidden'});});
 api.queueShiftOp({op:'delete',id:'a',month:'2026-09',by:'松井',sent:false});assert(!api.shiftEntriesFor('2026-09').some(x=>x.id==='a'));await api.runShiftOps();
 assert.equal(api.state.shiftOps.length,0);assert(api.shiftEntriesFor('2026-09').some(x=>x.id==='a'));
 // An unconfirmed add waits; a fresh month read that shows it settles it, one that does not re-sends it.
 api.setSend(async()=>{throw Object.assign(new Error('offline'),{code:'result_unconfirmed'});});
 api.queueShiftOp({op:'add',id:'n2',entry:e('n2','休み','鈴木','2026-09-29'),by:'鈴木',sent:false});await api.runShiftOps();assert.equal(api.state.shiftOps[0].unconfirmed,true);
 api.state.shiftData['2026-09'].entries.push(e('n2','休み','鈴木','2026-09-29'));api.reconcileShiftOps('2026-09');assert.equal(api.state.shiftOps.length,0);
 api.queueShiftOp({op:'add',id:'n3',entry:e('n3','休み','鈴木','2026-09-30'),by:'鈴木',sent:false});await api.runShiftOps();assert.equal(api.state.shiftOps[0].unconfirmed,true);
 sent=[];api.setSend(async p=>{sent.push(p);return {ok:true,id:p.id};});api.reconcileShiftOps('2026-09');await api.runShiftOps();assert.equal(sent.length,1);assert.equal(sent[0].id,'n3');assert.equal(api.state.shiftOps.length,0);
 console.log('PASS: notices need admin or token with date window and cache invalidation, シフト sheet creation, validation, duplicate id, display-format normalization, month filter and cache, own/admin delete with history, POST-only writes, month grid, holidays, chip labels, next/today shift, background add/delete, rejected and unconfirmed ops');
})().catch(e=>{console.error(e);process.exitCode=1});
