process.env.TZ='Asia/Tokyo';
const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),crypto=require('node:crypto');
const H=3600000,NOW=Date.now(),pad=n=>String(n).padStart(2,'0');
function jst(ms){const d=new Date(ms+9*H);return `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;}
const midnight=(()=>{const d=new Date(NOW+9*H);return Date.UTC(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())-9*H;})();
const at=(daysAgo,h,m=0)=>jst(midnight-daysAgo*24*H+h*H+m*60000);
// ---- Server twin, for parity of the fit check ----
const gas={Date,JSON,Math,console};vm.createContext(gas);vm.runInContext(fs.readFileSync('Code.gs','utf8'),gas);
// ---- Client with a small DOM ----
const nodes=new Map();
function el(id){ if(!nodes.has(id)){ const n={id,value:'',textContent:'',innerHTML:'',disabled:false,min:'',max:'',dataset:{},classes:new Set(),attrs:{},
  classList:{add(v){n.classes.add(v)},remove(v){n.classes.delete(v)},toggle(v,on){if(on===undefined) on=!n.classes.has(v);on?n.classes.add(v):n.classes.delete(v);},contains(v){return n.classes.has(v)}},
  setAttribute(k,v){n.attrs[k]=v},addEventListener(){},querySelectorAll:()=>[],insertAdjacentHTML(p,h){n.innerHTML+=h},focus(){},appendChild(){}}; nodes.set(id,n);} return nodes.get(id); }
let tmp=0;
const client={Date,URL,URLSearchParams,AbortController,Set,Map,console,setTimeout:()=>1,clearTimeout(){},setInterval(){},localStorage:{getItem:()=>null,setItem(){}},window:{addEventListener(){},crypto},document:{getElementById:el,querySelectorAll:()=>[],createElement:()=>el('__tmp'+(++tmp))}};
vm.createContext(client);
const script=fs.readFileSync('index.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
vm.runInContext(script.replace("if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();",`window.api={state,correctionFit,unclosedShift,openFixModal,fixCheck,submitFix,triggerPunchConfirmation,executePunch,renderRecentLogs,renderPunch,calculateMonthlyStats,summarizeMonth,renderEvents,openShiftModal,resumePendingPunch,punchQueueIdle,setSend:fn=>{adminRequest=fn;},setRead:fn=>{readGASDataOnce=fn;},prepare:()=>{renderAll=()=>{};showStatus=t=>{el('statusMessage').textContent=t;};}};`),client);
const api=client.window.api;api.prepare();
const tick=()=>new Promise(r=>setImmediate(r));
(async()=>{
 // Page and server agree on every verdict and message.
 { const L=[{id:'a',type:'出勤',time:at(5,9)},{id:'b',type:'退勤',time:at(5,15)},{id:'c',type:'出勤',time:at(6,9)},{id:'d',type:'休憩開始',time:at(6,12)}];
   const cases=[['退勤',at(6,16)],['休憩終了',at(6,12,30)],['退勤',at(6,11)],['出勤',at(5,10)],['退勤',at(7,10)],['退勤',jst(new Date(at(6,9).replace(' ','T')+'+09:00').getTime()+19*H)],['休憩終了',at(5,12)],['休憩開始',at(6,12,10)],['退勤',at(5,15)],['休憩開始',at(5,12)]];
   for(const who of ['橋本','田中']){ const logs=L.map(l=>Object.assign({name:who},l));
     for(const [type,time] of cases){ const server=gas.correctionFit_(logs,who,type,time,who==='田中'?'業務配分':'');const page=api.correctionFit(logs,who,type,time);
       assert.equal(page,server?server.message:null,`${who} ${type} ${time}`); } } }
 api.state.gasUrl='https://example.test/exec';api.state.users=['橋本','田中'];
 api.setRead(async u=>({ok:true,scope:'month',month:String(u).match(/month=([\d-]+)/)[1],logs:[]}));
 const sent=[];api.setSend(async p=>{sent.push(p);return {ok:true,action:'add',id:p.requestId,memo:'【修正】'+(p.reason||'押し忘れ')+'（入力）',allocations:JSON.parse(p.allocations||'[]'),corrected:p.correction==='1'?'修正 '+jst(Date.now()):''};});
 // ---- A forgotten clock-out is spotted, entered with its real time, and marked ----
 api.state.selectedUser='橋本';api.state.logs=[{id:'h1',name:'橋本',type:'出勤',time:at(2,8,30),month:at(2,8,30).slice(0,7),category:'',transport:'',memo:'',allocations:[]}];
 assert.equal(api.unclosedShift('橋本'),at(2,8,30));
 api.renderPunch();assert(!el('fixAlert').classes.has('hidden'));assert.match(el('fixAlertText').textContent,/出勤に退勤の記録がありません/);assert.equal(el('fixAlertBtn').dataset.date,at(2,8,30).slice(0,10));assert(!el('fixOpen').classes.has('hidden'));
 api.openFixModal();assert(el('fixModal').classes.has('show'));assert.equal(api.state.fixDraft.type,'退勤');assert.equal(api.state.fixDraft.date,at(2,8,30).slice(0,10));
 assert.match(api.fixCheck().preview,/時刻を入力/);await tick();
 api.state.fixDraft.time='15:30';let c=api.fixCheck();assert.equal(c.ok,true);assert.match(c.preview,/実働 7時間0分/);
 api.state.fixDraft.time='08:00';assert.match(api.fixCheck().error,/出勤の記録がありません/);
 api.state.fixDraft.time='15:30';el('fixReason').value='押し忘れました';api.submitFix();
 assert(!el('fixModal').classes.has('show'));const draft=api.state.confirmPunch;assert.equal(draft.correction.reason,'押し忘れました');assert.equal(draft.fixedTime.full,`${at(2,8,30).slice(0,10)} 15:30:00`);
 assert.match(el('confirmDetails').innerHTML,/fix-banner/);assert.match(el('confirmDetails').innerHTML,/赤字で残ります/);assert.equal(el('executePunch').textContent,'修正依頼を送る');
 el('checkoutMemo').value='';await api.executePunch();await api.punchQueueIdle();
 const p=sent[0];assert.equal(p.correction,'1');assert.equal(p.reason,'押し忘れました');assert.equal(p.time,`${at(2,8,30).slice(0,10)} 15:30:00`);assert.equal(p.type,'退勤');
 const saved=api.state.logs.find(l=>l.id===p.requestId);assert.match(saved.corrected,/^修正 /);assert.match(saved.memo,/^【修正】押し忘れました/);
 assert.equal(api.unclosedShift('橋本'),'','the missing clock-out is filled');
 api.renderRecentLogs();assert.match(el('recentLogs').innerHTML,/log-item[^"]* corrected/);assert.match(el('recentLogs').innerHTML,/✏️ 修正/);
 { const d=new Date(p.time.replace(' ','T')+'+09:00');assert(el('recentLogs').innerHTML.includes(`${d.getMonth()+1}/${d.getDate()} 15:30`),'a corrected punch shows its date'); }
 // The month counts it.
 { const m=p.time.slice(0,7);const st=api.calculateMonthlyStats(m);assert.equal(st['橋本'].corrections,1);assert.equal(api.summarizeMonth(st,['橋本','田中'],{}).totalCorrections,1); }
 // ---- 業務配分: the corrected checkout asks for the split of the hours up to that time ----
 api.state.selectedUser='田中';api.state.logs.push({id:'t1',name:'田中',type:'出勤',time:at(3,10),month:at(3,10).slice(0,7),category:'業務配分',transport:'',memo:'',allocations:[]});
 api.openFixModal({type:'退勤'});await tick();api.state.fixDraft.time='13:00';el('fixReason').value='';api.submitFix();
 assert.equal(api.state.confirmPunch.targetMinutes,180);assert.match(el('confirmDetails').innerHTML,/入力した退勤時刻までの実働時間/);api.state.confirmPunch=null;
 // A refused correction reopens its own form, ready to replace it.
 { const log={id:'r1',name:'橋本',type:'退勤',time:at(9,15),month:at(9,15).slice(0,7),category:'',transport:'',memo:'',allocations:[],corrected:'修正依頼',fixReason:'理由'};
   api.state.pendingPunch={gasUrl:api.state.gasUrl,log,receipt:crypto.randomUUID(),sent:true,error:'その時刻に出勤の記録がありません。',rejected:true};
   el('fixModal').classes.delete('show');api.resumePendingPunch();assert(el('fixModal').classes.has('show'));assert.equal(api.state.fixDraft.replaces,'r1');assert.equal(api.state.fixDraft.time,'15:00');
   assert.equal(el('fixError').textContent,'その時刻に出勤の記録がありません。');assert.equal(el('fixReason').value,'理由');api.state.pendingPunch=null; }
 // ---- 📌 予定・連絡: always on the punch screen with ＋ 追加; the calendar adds people only ----
 api.state.shiftData={};api.renderEvents();assert(!el('eventCard').classes.has('hidden'));assert.match(el('eventList').innerHTML,/ありません/);assert.match(el('eventSummary').textContent,/予定なし/);
 { const today=jst(Date.now()).slice(0,10);api.state.shiftData[today.slice(0,7)]={entries:[{id:'ev1',kind:'業務',name:'',start:today,end:today,from:'',to:'',label:'バイブル到着',memo:'',by:'',at:''}],changes:[]};
   api.renderEvents();assert.match(el('eventList').innerHTML,/バイブル到着/);assert.match(el('eventList').innerHTML,/data-event-edit="ev1"/);assert.match(el('eventList').innerHTML,/data-event-delete="ev1"/); }
 api.openShiftModal(jst(Date.now()).slice(0,10),null,{kinds:['業務']});assert.equal(JSON.stringify(api.state.shiftKinds),'["業務"]');assert(el('shiftKindRow').classes.has('hidden'));assert.equal(el('shiftModalTitle').textContent,'予定・連絡を追加');assert.equal(api.state.shiftKind,'業務');
 api.openShiftModal(jst(Date.now()).slice(0,10),null);assert(!api.state.shiftKinds.includes('業務'));assert(!el('shiftKindRow').classes.has('hidden'));assert.equal(api.state.shiftKind,'シフト');assert.doesNotMatch(el('shiftKindChoices').innerHTML,/予定\(終日\)/);
 api.openShiftModal(jst(Date.now()).slice(0,10),{id:'ev1',kind:'業務',name:'',start:'2026-09-30',end:'2026-09-30',label:'x',memo:''});assert.equal(el('shiftModalTitle').textContent,'予定・連絡を変更');
 console.log('test-corrections-client: OK');
})().catch(e=>{console.error(e);process.exit(1);});
