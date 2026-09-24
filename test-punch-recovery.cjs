const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),crypto=require('node:crypto');
const source=fs.readFileSync('index.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
function client(storage=new Map()){
 let now=0;class Clock extends Date {static now(){return now;}}
 const nodes=new Map();const el=id=>{if(!nodes.has(id)){const classes=new Set();nodes.set(id,{value:'',textContent:'',innerHTML:'',disabled:false,classes,classList:{toggle(k,v){v?classes.add(k):classes.delete(k)},add:k=>classes.add(k),remove:k=>classes.delete(k)},querySelectorAll:()=>[],insertAdjacentHTML(_,html){this.innerHTML+=html;}});}return nodes.get(id);};
 const ctx={Date:Clock,URL,URLSearchParams,AbortController,console,setTimeout:(fn,ms)=>{const t=setTimeout(fn,ms);if(t.unref)t.unref();return t;},clearTimeout,localStorage:{getItem:k=>storage.get(k)??null,setItem:(k,v)=>storage.set(k,v)},window:{location:{hostname:'hw-kintai-v2.vercel.app'},addEventListener(){},crypto},document:{getElementById:el,querySelectorAll:()=>[]}};
 vm.createContext(ctx);vm.runInContext(source.replace("if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();",`window.test={state,adminRequest,readGASDataOnce,executePunch,closePunchModal,resumePendingPunch,performSync,rememberPendingPunch,renderPendingPunch,readAllocationInputs,punchQueueIdle,drainPunchQueue,discardPendingPunch,
 setRead:fn=>readGASViaJSONP=fn,setSend:fn=>adminRequest=fn,setWait:fn=>wait=fn,setSyncRead:fn=>readGASDataOnce=fn,
 prepare:()=>{renderAll=()=>{};renderPunch=()=>{};setSyncing=()=>{};}};`),ctx);
 const api=ctx.window.test;api.prepare();api.setWait(async ms=>{now+=ms;});return {ctx,api,el,storage,advance:ms=>now+=ms};
}
const log={id:crypto.randomUUID(),name:'検証専用',type:'出勤',time:'2026-09-18 09:00:00',month:'2026-09',transport:'自転車・その他（支給なし）',memo:'固定メモ',category:'',allocations:[]};
const params={...log,requestId:log.id,action:'add',allocations:'[]',receipt:crypto.randomUUID()};
const draft=()=>({userName:log.name,punchType:log.type,transport:log.transport,memo:log.memo,category:'',fixedTime:{full:log.time,monthOnly:log.month,display:'09:00'}});
(async()=>{
 // A pending POST redirect must not block the confirmed receipt or cause a second POST.
 let c=client(),posts=[];c.ctx.fetch=(url,options)=>{posts.push({url,options});return new Promise(()=>{});};let reads=0;
 c.api.setRead(async url=>{assert.equal(new URL(url).searchParams.get('receipt'),params.receipt);reads++;return reads===1?{ok:false,pending:true}:{ok:true,action:'add',id:log.id};});
 assert.equal((await c.api.adminRequest(params)).id,log.id);assert.equal(posts.length,1);assert.equal(posts[0].options.method,'POST');assert.equal(posts[0].options.mode,'no-cors');assert.equal(posts[0].options.credentials,'omit');assert.equal(posts[0].options.body.get('requestId'),log.id);assert.equal(posts[0].options.body.get('receipt'),params.receipt);assert.equal(posts[0].options.signal.aborted,true);
 // Default reads immediately use Google, with no 25-second API detour.
 c.api.setRead(async url=>{assert(url.startsWith('https://script.google.com/'));return {ok:true,scope:'recent',logs:[]};});assert.equal((await c.api.readGASDataOnce(posts[0].url+'?action=read&scope=recent')).ok,true);assert.equal(posts.length,1);
 // An opaque successful response is not a save confirmation. Timeouts never replay the write.
 c=client();let sent=0;c.ctx.fetch=async()=>{sent++;return {type:'opaque'};};c.api.setRead(async()=>{c.advance(1500);return {ok:false,pending:true};});
 await assert.rejects(()=>c.api.adminRequest(params),e=>e.code==='result_unconfirmed');assert.equal(sent,1);
 // A failed POST response may still have saved. The receipt is authoritative.
 c=client();c.ctx.fetch=async()=>{throw Error('network lost');};c.api.setRead(async()=>({ok:true,action:'add',id:log.id}));assert.equal((await c.api.adminRequest(params)).id,log.id);
 c.api.setRead(async()=>({ok:false,error:'allocation_mismatch',actualMinutes:165,message:'配分を確認'}));await assert.rejects(()=>c.api.adminRequest(params),e=>e.code==='allocation_mismatch'&&e.actualMinutes===165);
 // Ambiguous write -> close modal -> reload -> receipt expired -> locate saved ID without writing again.
 c=client();c.api.state.confirmPunch=draft();let payload;
 c.api.setSend(async p=>{payload=p;assert(JSON.parse(c.storage.get('attendance_pending_punch')).log.id===p.id);throw Object.assign(Error('offline'),{code:'result_unconfirmed'});});
 await c.api.executePunch();await c.api.punchQueueIdle();assert.equal(c.api.state.isPunchSubmitting,false);assert(c.api.state.pendingPunch);assert.equal(c.api.state.pendingPunch.sent,true);assert.match(c.api.state.pendingPunch.error,/offline/);assert.equal(c.api.state.logs.length,0);assert.match(c.el('statusMessage').textContent,/保存確認・再送/);
 c.api.closePunchModal();assert(c.api.state.pendingPunch);const pending=c.api.state.pendingPunch;
 const reloaded=client(c.storage);reloaded.api.resumePendingPunch();assert.equal(reloaded.api.state.confirmPunch.log.id,pending.log.id);assert.equal(reloaded.api.state.confirmPunch.fixedTime.full,pending.log.time);
 let recoveryWrites=0;reloaded.api.setSend(async()=>{recoveryWrites++;throw Error('must not write');});
 reloaded.api.setRead(async url=>new URL(url).searchParams.get('action')==='adminResult'?{ok:false,pending:true}:{ok:true,scope:'month',month:pending.log.month,logs:[pending.log]});
 await reloaded.api.executePunch();assert.equal(recoveryWrites,0);assert.equal(reloaded.api.state.pendingPunch,null);assert.equal(reloaded.api.state.logs[0].id,pending.log.id);assert.equal(reloaded.api.state.confirmPunch,null);
 // Missing ID permits one explicit retry with the original ID/time/payload; receipt is renewed.
 c=client();c.api.rememberPendingPunch(log,params.receipt);c.api.resumePendingPunch();c.api.setRead(async url=>new URL(url).searchParams.get('action')==='adminResult'?{ok:false,pending:true}:{ok:true,scope:'month',month:log.month,logs:[]});
 let release,count=0;c.api.setSend(p=>{count++;assert.equal(p.id,log.id);assert.equal(p.time,log.time);assert.equal(p.memo,log.memo);assert.notEqual(p.receipt,params.receipt);return new Promise(r=>release=r);});
 const first=c.api.executePunch();for(let i=0;i<12&&!release;i++)await Promise.resolve();await c.api.executePunch();assert.equal(count,1);release({ok:true,id:log.id});await first;assert.equal(c.api.state.pendingPunch,null);
 // A fresh background read confirms recovery; an old in-flight read cannot undo a newer save.
 c=client();c.api.rememberPendingPunch(log,params.receipt);c.api.setSyncRead(async()=>({ok:true,scope:'recent',logs:[log]}));await c.api.performSync('auto');assert.equal(c.api.state.pendingPunch,null);
 let finish;c.api.setSyncRead(()=>new Promise(r=>finish=r));const sync=c.api.performSync('auto');c.api.state.mutationVersion++;finish({ok:true,scope:'recent',logs:[]});await sync;assert.equal(c.api.state.logs[0].id,log.id);
 // Reloaded checkout retains editable allocation fields if validation is rejected, and its original time.
 c=client();const ids=['hw_production','hw_support','hw_admin','hw_pro','hw_other','hwa_production','hwa_support','hwa_admin','hwa_text','hwa_other'];const checkout={...log,type:'退勤',name:'田中',category:'業務配分',allocations:ids.map((id,i)=>({id,minutes:i?0:60,memo:i?'':'作業'}))};
 c.api.rememberPendingPunch(checkout,params.receipt);c.api.resumePendingPunch();assert.equal(c.api.readAllocationInputs()[0].minutes,60);assert.match(c.el('confirmDetails').innerHTML,/checkoutMemo/);
 c.api.setRead(async()=>({ok:false,error:'allocation_mismatch',actualMinutes:65,message:'配分を確認'}));await c.api.executePunch();assert.equal(c.api.state.pendingPunch,null);assert.equal(c.api.state.confirmPunch.log,null);assert.equal(c.api.state.confirmPunch.targetMinutes,65);assert.equal(c.api.state.confirmPunch.fixedTime.full,log.time);
 // A first send closes the dialog at once, goes out in the background, and confirms without reopening anything.
 c=client();c.api.state.confirmPunch=draft();c.el('confirmModal').classList.add('show');let settle;c.api.setSend(()=>new Promise(r=>settle=r));
 await c.api.executePunch();assert.equal(c.el('confirmModal').classes.has('show'),false);assert.equal(c.api.state.isPunchSubmitting,false);assert.equal(c.api.state.confirmPunch,null);assert.equal(c.api.state.pendingPunch.sent,true);assert.equal(c.api.state.logs.length,0);assert.match(c.el('statusMessage').textContent,/記録しました/);
 settle({ok:true,action:'add',id:c.api.state.pendingPunch.log.id});await c.api.punchQueueIdle();assert.equal(c.api.state.pendingPunch,null);assert.equal(c.api.state.logs.length,1);assert.equal(c.el('confirmModal').classes.has('show'),false);
 // A second punch while the first is still confirming is queued, then sent by itself in order.
 c=client();const resolvers=[];c.api.setSend(()=>new Promise(r=>resolvers.push(r)));
 c.api.state.confirmPunch=draft();await c.api.executePunch();const firstId=c.api.state.pendingPunch.log.id;
 c.api.state.confirmPunch={...draft(),punchType:'休憩開始',fixedTime:{full:'2026-09-18 10:00:00',monthOnly:'2026-09',display:'10:00'}};await c.api.executePunch();
 assert.equal(resolvers.length,1);assert.equal(c.api.state.punchQueue.length,1);assert.equal(c.api.state.pendingPunch.log.id,firstId);assert.equal(JSON.parse(c.storage.get('attendance_punch_queue')).length,1);
 let idle=c.api.punchQueueIdle();resolvers[0]({ok:true,action:'add',id:firstId});await idle;assert.equal(resolvers.length,2);assert.equal(c.api.state.punchQueue.length,0);assert.equal(c.api.state.pendingPunch.log.type,'休憩開始');assert.equal(c.api.state.logs.length,1);
 idle=c.api.punchQueueIdle();resolvers[1]({ok:true,action:'add',id:c.api.state.pendingPunch.log.id});await idle;assert.equal(c.api.state.pendingPunch,null);assert.equal(c.api.state.logs.length,2);assert.equal(c.api.state.logs[0].type,'休憩開始');
 // A rejected send is shown on the card, not in a surprise dialog; "修正して再送" reopens it editable at the original time.
 c=client();c.api.state.confirmPunch=draft();c.api.setSend(async()=>{throw Object.assign(Error('配分を確認'),{code:'allocation_mismatch',actualMinutes:65});});
 await c.api.executePunch();await c.api.punchQueueIdle();assert.equal(c.el('confirmModal').classes.has('show'),false);assert.equal(c.api.state.pendingPunch.rejected,true);assert.equal(c.api.state.pendingPunch.actualMinutes,65);assert.match(c.el('statusMessage').textContent,/配分を確認/);
 const rejectedId=c.api.state.pendingPunch.log.id;c.api.resumePendingPunch();assert.equal(c.el('confirmModal').classes.has('show'),true);assert.equal(c.api.state.confirmPunch.log,null);assert.equal(c.api.state.confirmPunch.retry,false);assert.equal(c.api.state.confirmPunch.replaces,rejectedId);assert.equal(c.api.state.confirmPunch.targetMinutes,65);assert.equal(c.el('executePunch').textContent,'はい、打刻する');
 c.api.setSend(async p=>({ok:true,action:'add',id:p.id}));await c.api.executePunch();await c.api.punchQueueIdle();assert.equal(c.api.state.pendingPunch,null);assert.equal(c.api.state.punchQueue.length,0);assert.equal(c.api.state.logs.length,1);assert.equal(c.api.state.logs[0].time,log.time);assert.notEqual(c.api.state.logs[0].id,rejectedId);
 // An unconfirmed send stays closed and hands over to the pending card; discarding it clears the way.
 c=client();c.api.state.confirmPunch=draft();c.api.setSend(async()=>{throw Object.assign(Error('offline'),{code:'result_unconfirmed'});});
 await c.api.executePunch();await c.api.punchQueueIdle();assert.equal(c.el('confirmModal').classes.has('show'),false);assert(c.api.state.pendingPunch);assert.equal(c.api.state.pendingPunch.rejected,false);assert.equal(c.api.state.confirmPunch,null);assert.match(c.el('statusMessage').textContent,/保存確認・再送/);
 c.api.discardPendingPunch();assert.equal(c.api.state.pendingPunch,null);assert.equal(c.storage.get('attendance_pending_punch'),'null');
 // After a reload, an already-sent punch is never re-sent by itself, but a queued one continues once the head is confirmed.
 c=client();const held=[];c.api.setSend(()=>new Promise(r=>held.push(r)));c.api.state.confirmPunch=draft();await c.api.executePunch();const headId=c.api.state.pendingPunch.log.id;
 c.api.state.confirmPunch={...draft(),punchType:'休憩開始',fixedTime:{full:'2026-09-18 10:00:00',monthOnly:'2026-09',display:'10:00'}};await c.api.executePunch();assert.equal(c.api.state.punchQueue.length,1);
 const again=client(c.storage);assert.equal(again.api.state.pendingPunch.log.id,headId);assert.equal(again.api.state.pendingPunch.sent,true);assert.equal(again.api.state.punchQueue.length,1);
 let resent=0;again.api.setSend(async p=>{resent++;return {ok:true,action:'add',id:p.id};});again.api.drainPunchQueue();await again.api.punchQueueIdle();assert.equal(resent,0);assert(again.api.state.pendingPunch);
 again.api.setSyncRead(async()=>({ok:true,scope:'recent',logs:[again.api.state.pendingPunch.log]}));await again.api.performSync('startup');await again.api.punchQueueIdle();assert.equal(resent,1);assert.equal(again.api.state.pendingPunch,null);assert.equal(again.api.state.punchQueue.length,0);assert.equal(again.api.state.logs.length,2);
 // A retry keeps the dialog open until the outcome is known.
 c=client();c.api.rememberPendingPunch(log,params.receipt);c.api.resumePendingPunch();let hold;c.api.setRead(async url=>new URL(url).searchParams.get('action')==='adminResult'?{ok:false,pending:true}:{ok:true,scope:'month',month:log.month,logs:[]});c.api.setSend(()=>new Promise(r=>hold=r));
 const retry=c.api.executePunch();for(let i=0;i<12&&!hold;i++)await Promise.resolve();assert.equal(c.el('confirmModal').classes.has('show'),true);hold({ok:true,id:log.id});await retry;assert.equal(c.el('confirmModal').classes.has('show'),false);
 console.log('PASS: background first send, ordered queue, rejected send on card with editable resend, unconfirmed hand-over and discard, reload keeps queue without re-sending, retry waits, direct read, concurrent exact receipt, one POST, no false success, recovery after reload/cache expiry, explicit idempotent retry, immediate lock/release, preserved checkout fields/time, stale-read guard');
})().catch(e=>{console.error(e);process.exitCode=1});
