const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const rows=[['timestamp','name','type','time','month','id','transport','memo','category']];let serial=0;
const sheet={getLastRow:()=>rows.length,appendRow:r=>rows.push(r),getRange:(r,c,n=1,w=1)=>({createTextFinder(value){const found=rows.slice(r-1,r-1+n).flatMap((row,i)=>String(row[c-1])===value?[{getRow:()=>r+i}]:[]);return {matchEntireCell(){return this},findNext:()=>found[0]||null,findAll:()=>found}},getDisplayValues:()=>rows.slice(r-1,r-1+n).map(row=>row.slice(c-1,c-1+w)),setFormula(){return this},setNumberFormat(){return this}})};
const cache=new Map(),props=new Map();
const gas={CacheService:{getScriptCache:()=>({get:k=>cache.get(k)||null,put:(k,v)=>cache.set(k,v),remove:k=>cache.delete(k)})},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v),deleteProperty:k=>props.delete(k)})},Utilities:{getUuid:()=>require('crypto').randomUUID()},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},SpreadsheetApp:{flush(){}}};vm.createContext(gas);vm.runInContext(fs.readFileSync('Code.gs','utf8')+';this.items=ALLOCATION_ITEMS;',gas);
const punch=(type,time,extra={})=>gas.addLog_(sheet,{id:'id-'+(++serial),name:'田中',type,time:'2026-09-17 '+time+':00',month:'2026-09',category:'業務配分',...extra});
assert.equal(punch('出勤','09:00').ok,true);assert.equal(punch('休憩開始','10:00').ok,true);assert.equal(punch('休憩終了','10:15').ok,true);
const alloc=gas.items.map(i=>({id:i.id,minutes:0,memo:''}));alloc[0].minutes=60;alloc[0].memo='伝票作成';alloc[8].minutes=105;alloc[8].memo='動画制作';
assert.equal(punch('退勤','12:00',{allocations:'[]'}).error,'invalid_allocation');
const wrong=structuredClone(alloc);wrong[0].minutes=61;assert.equal(punch('退勤','12:00',{allocations:JSON.stringify(wrong)}).error,'allocation_mismatch');assert.equal(rows.length,4);
const result=punch('退勤','12:00',{allocations:JSON.stringify(alloc)});assert.equal(result.ok,true);assert.equal(result.memo,'HWの生産（60分）：伝票作成｜HWAのテキスト制作（105分）：動画制作');assert.equal(rows[4][10],60/1440);assert.equal(rows[4][18],105/1440);assert.equal(JSON.parse(rows[4][20]).length,10);
assert.equal(punch('退勤','12:00',{allocations:JSON.stringify(alloc)}).duplicate,true);assert.equal(rows.length,5);
assert.equal(punch('出勤','13:00',{name:'橋本'}).error,'invalid_category');
const duplicate=structuredClone(alloc);duplicate[1].id=duplicate[0].id;assert.equal(gas.validateAllocations_(duplicate).ok,false);
const bad=structuredClone(alloc);bad[0].minutes=-1;assert.equal(gas.validateAllocations_(bad).ok,false);
const logs=[{name:'田中',type:'出勤',time:'2026-09-17 23:00:00'},{name:'田中',type:'休憩開始',time:'2026-09-18 00:00:00'},{name:'田中',type:'休憩終了',time:'2026-09-18 00:30:00'}];assert.equal(gas.shiftMinutes_(logs,'田中','2026-09-18 02:00:00'),150);
// getDisplayValues() renders "09:22" as "9:22", which is not valid ISO 8601 and used to drop the clock-in.
assert.equal(gas.shiftMinutes_([{name:'牛嶋',type:'出勤',time:'2026-09-18 9:22:45'}],'牛嶋','2026-09-18 18:00:00'),517);
assert.equal(gas.shiftMinutes_([{name:'牛嶋',type:'出勤',time:'2026-09-20 0:03:07'}],'牛嶋','2026-09-20 09:00:00'),536);
assert.equal(gas.shiftMinutes_([{name:'牛嶋',type:'出勤',time:'2026/9/18 9:22:45'}],'牛嶋','2026-09-18 18:00:00'),517);
// A clock-out more than 18 hours after the clock-in belongs to no shift; a later clock-in starts a new one.
assert.equal(gas.shiftMinutes_([{name:'牛嶋',type:'出勤',time:'2026-09-17 09:00:00'}],'牛嶋','2026-09-18 09:00:00'),null);
assert.equal(gas.shiftMinutes_([{name:'牛嶋',type:'出勤',time:'2026-09-17 09:00:00'},{name:'牛嶋',type:'出勤',time:'2026-09-18 09:00:00'}],'牛嶋','2026-09-18 18:00:00'),540);
assert.equal(gas.shiftMinutes_([{name:'牛嶋',type:'出勤',time:'2026-09-17 09:00:00'},{name:'牛嶋',type:'休憩開始',time:'2026-09-18 10:00:00'}],'牛嶋','2026-09-18 18:00:00'),null);
{ // the read index closes a session that outlived the limit, so the next day's punches start clean
  const idx={version:1,lastRow:1,builtAt:Date.now(),months:{},latest:{},sessions:{},recent:[],maxTime:'',revision:'r'};
  gas.updateReadIndex_(idx,[['t','牛嶋','出勤','2026-09-17 09:00:00','2026-09','a','','','業務配分'],['t','牛嶋','退勤','2026-09-18 09:30:00','2026-09','b','','','業務配分'],['t','牛嶋','出勤','2026-09-18 10:00:00','2026-09','c','','','業務配分']],2);
  assert.equal(JSON.stringify(idx.sessions['牛嶋']),'[4]');assert.equal(idx.sessionStart['牛嶋'],'2026-09-18 10:00:00');
}
const html=fs.readFileSync('index.html','utf8'),script=html.match(/<script>([\s\S]*?)<\/script>/)[1];const nodes=new Map();
function el(id){if(!nodes.has(id))nodes.set(id,{value:'',textContent:'',innerHTML:'',disabled:false,querySelectorAll:()=>[],classList:{add(){},remove(){},toggle(){}},focus(){}});return nodes.get(id);}
const client={Date,URLSearchParams,Set,console,localStorage:{getItem(){return null},setItem(){}},window:{addEventListener(){},crypto:{randomUUID:()=> 'request-1'}},document:{getElementById:el,querySelectorAll:()=>[]},setTimeout(){return 1},clearTimeout(){}};vm.createContext(client);
vm.runInContext(script.replace("if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();",`window.test={state,triggerPunchConfirmation,executePunch,calculateMonthlyStats,readAllocationInputs,syncLogsFromGAS,applyGasReadData,punchQueueIdle,shiftStateOf,setSend:fn=>{adminRequest=fn;},setSync:fn=>{performSync=fn;},prepare:()=>{renderAll=()=>{};renderPunch=()=>{};getJSTDateTime=()=>({full:'2026-09-17 12:00:00',monthOnly:'2026-09',display:'12:00'});}};`),client);
const api=client.window.test;api.prepare();api.state.selectedUser='田中';api.state.users=['田中'];api.state.logs=rows.slice(1,4).map(r=>({id:r[5],name:r[1],type:r[2],time:r[3],month:r[4],category:r[8],transport:'',allocations:[]}));
api.triggerPunchConfirmation('退勤');assert.equal(api.state.confirmPunch.targetMinutes,165);assert.match(el('confirmDetails').innerHTML,/HWAのテキスト制作/);assert.equal(api.state.confirmPunch.category,'業務配分');
el('hw_production_hours').value='1';el('hw_production_memo').value='伝票作成';el('hwa_text_hours').value='1';el('hwa_text_minutes').value='45';el('hwa_text_memo').value='動画制作';
el('hw_support_minutes').value='-1';assert.equal(Number.isNaN(api.readAllocationInputs()[1].minutes),true);el('hw_support_minutes').value='';
assert.throws(()=>api.applyGasReadData({ok:false,error:'read_error'}),/read_error/);assert.throws(()=>api.applyGasReadData({ok:true}),/読み取れません/);
(async()=>{
 let sends=0,resolve;api.setSend(()=>{sends++;return new Promise(r=>resolve=r)});
 await api.executePunch();await api.executePunch();assert.equal(sends,1);assert.equal(api.state.isPunchSubmitting,false);assert.equal(api.state.confirmPunch,null);assert(api.state.pendingPunch);resolve({...result,id:'request-1'});await api.punchQueueIdle();assert.equal(api.state.pendingPunch,null);assert.equal(api.state.logs[0].memo,result.memo);
 const stats=api.calculateMonthlyStats('2026-09')['田中'];assert.equal(stats.totalMinutes,165);assert.equal(stats.categoryMinutes['ホームワイン'],60);assert.equal(stats.categoryMinutes['アカデミー'],105);
 let syncs=0,done;api.setSync(()=>{syncs++;return new Promise(r=>done=r)});const first=api.syncLogsFromGAS('manual'),second=api.syncLogsFromGAS('tab');assert.equal(first,second);assert.equal(syncs,1);done();await first;
 api.state.selectedUser='田中';api.state.confirmPunch=null;api.state.logs=[{id:'s1',name:'田中',type:'出勤',time:'2026-09-17 9:22:45',month:'2026-09',category:'業務配分',transport:'',memo:'',allocations:[]}];
 api.triggerPunchConfirmation('退勤');assert.equal(api.state.confirmPunch.targetMinutes,157);
 api.state.selectedUser='橋本';api.triggerPunchConfirmation('出勤');assert.equal(api.state.confirmPunch.category,'');
 api.state.confirmPunch=null;api.state.users=['橋本'];api.state.logs=[{id:'old',name:'橋本',type:'出勤',time:'2026-09-16 09:00:00',month:'2026-09',category:'',transport:'',memo:'',allocations:[]}];
 api.triggerPunchConfirmation('退勤');assert.equal(api.state.confirmPunch,null);assert.match(el('statusMessage').textContent,/18時間以上/);
 api.triggerPunchConfirmation('休憩開始');assert.equal(api.state.confirmPunch,null);
 api.state.logs=[];api.triggerPunchConfirmation('退勤');assert.match(el('statusMessage').textContent,/先に「出勤」/);
 api.state.logs=[{id:'today',name:'橋本',type:'出勤',time:'2026-09-17 09:00:00',month:'2026-09',category:'',transport:'',memo:'',allocations:[]}];api.triggerPunchConfirmation('退勤');assert.equal(api.state.confirmPunch.punchType,'退勤');assert.equal(api.state.confirmPunch.targetMinutes,null);
 { const stats=api.calculateMonthlyStats('2026-09'); assert.equal(stats['橋本'].totalMinutes,0); }
 // Buttons follow the shift: no second clock-in while one is open, breaks alternate, clock-in returns after checkout.
 const w=(id,type,time)=>({id,name:'橋本',type,time,month:'2026-09',category:'',transport:'',memo:'',allocations:[]});
 api.state.logs=[w('w1','出勤','2026-09-17 09:00:00')];assert.equal(api.shiftStateOf('橋本').state,'working');api.state.confirmPunch=null;
 api.triggerPunchConfirmation('出勤');assert.equal(api.state.confirmPunch,null);assert.match(el('statusMessage').textContent,/すでに出勤/);api.triggerPunchConfirmation('休憩終了');assert.equal(api.state.confirmPunch,null);
 api.state.logs.push(w('w2','休憩開始','2026-09-17 10:00:00'));assert.equal(api.shiftStateOf('橋本').state,'break');api.triggerPunchConfirmation('休憩開始');assert.equal(api.state.confirmPunch,null);api.triggerPunchConfirmation('退勤');assert.equal(api.state.confirmPunch.punchType,'退勤');
 api.state.logs.push(w('w3','退勤','2026-09-17 11:00:00'));assert.equal(api.shiftStateOf('橋本').state,'off');api.state.confirmPunch=null;api.triggerPunchConfirmation('出勤');assert.equal(api.state.confirmPunch.punchType,'出勤');
 assert(vm.runInContext('CATEGORY_USERS',gas).includes('鈴木'));api.state.users=['鈴木'];api.state.selectedUser='鈴木';api.state.confirmPunch=null;api.triggerPunchConfirmation('出勤');assert.equal(api.state.confirmPunch.category,'業務配分');
 console.log('PASS: 10 allocations, net shift time, overnight break, unpadded sheet hours, 18-hour shift limit, shift-state buttons, 鈴木 in work categories, mismatch/negative/duplicate rejection, pipe memo, sheet numeric durations, queued checkout confirmed in background, monthly categories, shared sync promise and failed response rejection');
})().catch(e=>{console.error(e);process.exitCode=1});
