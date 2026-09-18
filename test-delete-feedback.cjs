const fs=require('fs'),vm=require('vm'),assert=require('assert/strict'),crypto=require('crypto');
const nodes=new Map();function el(id){if(!nodes.has(id))nodes.set(id,{value:'',textContent:'',innerHTML:'',disabled:false,attrs:{},classes:new Set(),classList:{add(v){nodes.get(id).classes.add(v)},remove(v){nodes.get(id).classes.delete(v)},toggle(v,on){on?nodes.get(id).classes.add(v):nodes.get(id).classes.delete(v)}},setAttribute(k,v){this.attrs[k]=v},focus(){},querySelectorAll:()=>[]});return nodes.get(id)}
const client={Date,URL,URLSearchParams,AbortController,console,setTimeout:()=>1,clearTimeout(){},localStorage:{getItem:()=>null,setItem(){}},window:{addEventListener(){}},document:{getElementById:el,querySelectorAll:()=>[]}};
vm.createContext(client);const script=fs.readFileSync('index.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];vm.runInContext(script.replace("if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();",`window.api={state,requestDeleteLog,confirmDeleteLog,closeDeleteModal,send:fn=>adminRequest=fn,read:fn=>readGASDataOnce=fn,prepare:()=>{renderAll=()=>{};showStatus=()=>{};}};`),client);
const api=client.window.api;api.prepare();const log={id:'delete-test',name:'検証用',type:'出勤',time:'2026-09-18 09:00:00',month:'2026-09'};
function open(){api.state.logs=[log];api.state.monthLogs=[log];api.requestDeleteLog(log.id);el('deletePassword').value='test-only';}
(async()=>{
 open();let calls=0,resolve;api.send(p=>{calls++;assert.equal(p.action,'delete');assert.equal(p.adminPassword,'test-only');return new Promise(r=>resolve=r)});
 const first=api.confirmDeleteLog();assert.equal(el('executeDelete').textContent,'取り消し中…');assert.equal(el('executeDelete').disabled,true);assert.equal(el('cancelDelete').disabled,true);assert.equal(el('deletePassword').value,'');assert.match(el('deleteProgress').textContent,/取り消し/);
 await api.confirmDeleteLog();api.closeDeleteModal();assert.equal(calls,1);assert(el('deleteModal').classes.has('show'));resolve({ok:true});await first;assert(!el('deleteModal').classes.has('show'));assert.equal(api.state.logs.length,0);assert.equal(api.state.monthLogs.length,0);assert.equal(api.state.isAdminSubmitting,false);
 open();api.send(async()=>{throw Object.assign(new Error('パスワードが正しくありません。'),{code:'unauthorized'})});await api.confirmDeleteLog();assert(el('deleteModal').classes.has('show'));assert.equal(el('deleteError').classes.has('hidden'),false);assert.match(el('deleteError').textContent,/パスワード/);assert.equal(el('executeDelete').disabled,false);assert.equal(api.state.logs.length,1);assert.equal(api.state.deleteNeedsCheck,false);
 open();api.send(async()=>{throw Object.assign(new Error('未確認'),{code:'result_unconfirmed'})});await api.confirmDeleteLog();assert.equal(el('executeDelete').textContent,'最新状態を確認');assert.equal(api.state.logs.length,1);assert(el('deleteModal').classes.has('show'));
 api.read(async()=>({ok:true,scope:'month',month:log.month,logs:[log]}));await api.confirmDeleteLog();assert.match(el('deleteProgress').textContent,/残っています/);assert.equal(api.state.deleteNeedsCheck,false);assert.equal(el('deletePassword').disabled,false);
 el('deletePassword').value='test-only';await api.confirmDeleteLog();api.read(async()=>({ok:true,scope:'month',month:log.month,logs:[]}));await api.confirmDeleteLog();assert.equal(api.state.logs.length,0);assert(!el('deleteModal').classes.has('show'));
 // Recent-list confirmation requires no password; history still does.
 open();api.requestDeleteLog(log.id,'recent');assert(el('deletePasswordBox').classes.has('hidden'));assert.equal(el('deletePassword').value,'');
 let recentCalls=0;api.send(async p=>{recentCalls++;assert.equal(p.action,'deleteRecent');assert.equal(p.id,log.id);assert.equal('adminPassword' in p,false);return {ok:true}});
 assert(el('deleteModal').classes.has('show'));assert.equal(recentCalls,0);await api.confirmDeleteLog();assert.equal(recentCalls,1);assert(!el('deleteModal').classes.has('show'));
 open();api.requestDeleteLog(log.id);assert(!el('deletePasswordBox').classes.has('hidden'));await api.confirmDeleteLog();assert.equal(recentCalls,1);assert.match(el('deleteError').textContent,/パスワードを入力/);
 open();api.requestDeleteLog(log.id,'recent');api.send(async()=>{throw Object.assign(new Error('未確認'),{code:'result_unconfirmed'})});await api.confirmDeleteLog();
 api.read(async()=>({ok:true,scope:'month',month:log.month,logs:[log]}));await api.confirmDeleteLog();assert.doesNotMatch(el('deleteProgress').textContent,/パスワード/);assert.equal(api.state.deleteNeedsCheck,false);
 api.send(async p=>{assert.equal(p.action,'deleteRecent');return {ok:true}});await api.confirmDeleteLog();assert.equal(api.state.logs.length,0);
 // Server combines authentication and deletion in one POST; invalid passwords cannot mutate rows.
 const cache=new Map(),props=new Map([['ADMIN_PASSWORD_SHA256',crypto.createHash('sha256').update('test-only').digest('hex')]]);let deletes=0,scans=0;
 const rows=[['timestamp','name','type','time','month','id'],['','検証用','出勤',log.time,log.month,log.id]];
 const sheet={getLastRow:()=>rows.length,deleteRow:n=>{rows.splice(n-1,1);deletes++;},getDataRange(){scans++;throw Error('full scan forbidden');},getRange:(r,c,n,w)=>({createTextFinder:value=>{const found=rows.slice(r-1,r-1+n).flatMap((row,i)=>String(row[c-1])===value?[{getRow:()=>r+i}]:[]);return {matchEntireCell(){return this},matchCase(){return this},findNext:()=>found[0]||null,findAll:()=>found}},getDisplayValues:()=>rows.slice(r-1,r-1+n).map(row=>row.slice(c-1,c-1+w))})};
 const gas={CacheService:{getScriptCache:()=>({get:k=>cache.get(k)||null,put:(k,v)=>cache.set(k,v),remove:k=>cache.delete(k)})},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,deleteProperty:k=>props.delete(k)})},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},Utilities:{getUuid:()=>crypto.randomUUID(),DigestAlgorithm:{SHA_256:''},Charset:{UTF_8:''},computeDigest:(_,v)=>[...crypto.createHash('sha256').update(v).digest()]},SpreadsheetApp:{openById:()=>({}),flush(){}},ContentService:{MimeType:{JSON:'json'},createTextOutput:value=>({setMimeType:()=>JSON.parse(value)})}};
 vm.createContext(gas);vm.runInContext(fs.readFileSync('Code.gs','utf8'),gas);gas.getOrCreateMasterSheet_=()=>sheet;
 const request=password=>({parameter:{action:'delete',id:log.id,adminPassword:password,receipt:crypto.randomUUID()}});
 assert.equal(gas.doPost(request('wrong')).error,'unauthorized');assert.equal(deletes,0);
 assert.equal(gas.doPost(request('test-only')).ok,true);assert.equal(deletes,1);assert.equal(scans,0);assert.equal(gas.doPost(request('test-only')).error,'not_found');assert.equal(deletes,1);
 // Unauthenticated cancellation is POST-only and limited to the displayed eight records.
 for(let i=0;i<10;i++) rows.push(['','検証用','出勤','2026-09-18 '+String(i+1).padStart(2,'0')+':00:00','2026-09','recent-'+i]);
 gas.getReadIndex_=()=>({recent:rows.slice(1).map((_,i)=>i+2)});
 const recentRequest=id=>({parameter:{action:'deleteRecent',id,receipt:crypto.randomUUID()}});
 assert.equal(gas.doGet(recentRequest('recent-9')).error,'post_required');
 const badReceipt=recentRequest('recent-9');badReceipt.parameter.receipt='invalid';assert.equal(gas.doPost(badReceipt).error,'invalid_receipt');
 assert.equal(gas.doPost(recentRequest('recent-1')).error,'not_recent');assert.equal(deletes,1);
 assert.equal(gas.doPost(recentRequest('recent-2')).ok,true);assert.equal(deletes,2);
 assert.equal(gas.doPost(recentRequest('recent-9')).ok,true);assert.equal(deletes,3);
 assert.equal(gas.doPost(recentRequest('recent-9')).error,'not_recent');assert.equal(deletes,3);
 assert.equal(gas.doPost(recentRequest('')).error,'missing_delete_params');
 assert.equal(gas.doPost({parameter:{action:'delete',id:'recent-8',receipt:crypto.randomUUID()}}).error,'unauthorized');
 assert.equal(scans,0);
 console.log('PASS: immediate feedback, single submission, close guard, success close/remove, visible auth errors, uncertain outcome verification, lock release, passwordless recent confirmation, protected history, POST-only recent scope, no full-sheet scan or duplicate deletion');
})().catch(e=>{console.error(e);process.exitCode=1});
