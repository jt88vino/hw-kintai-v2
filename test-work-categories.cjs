const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const html = fs.readFileSync(__dirname+'/index.html','utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const nodes = new Map();
function el(id) { if(!nodes.has(id)) nodes.set(id,{value:'',innerHTML:'',textContent:'',disabled:false,classList:{add(){},remove(){},toggle(){}},focus(){}}); return nodes.get(id); }
const client={console,Date,URLSearchParams,Set,localStorage:{getItem(){return null},setItem(){}},window:{addEventListener(){},crypto:{randomUUID:()=> 'request-1'}},document:{getElementById:el,querySelectorAll:()=>[]},setTimeout(){return 1},clearTimeout(){}};
vm.createContext(client);
vm.runInContext(script.replace("if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();","window.test = {state,sanitizeLog,calculateMonthlyStats,triggerPunchConfirmation,executePunch,setRender:()=>{renderAll=()=>{};renderPunch=()=>{};},setSend:fn=>{readGASViaJSONP=fn;}};"),client);
const api=client.window.test;
api.state.selectedUser='田中'; api.state.workCategory=''; api.triggerPunchConfirmation('出勤'); assert.equal(api.state.confirmPunch,null);
api.state.workCategory='アカデミー'; api.triggerPunchConfirmation('退勤'); assert.match(el('confirmDetails').innerHTML,/checkoutMemo/); assert.equal(api.state.confirmPunch.category,'アカデミー');
api.state.selectedUser='橋本'; api.triggerPunchConfirmation('出勤'); assert.equal(api.state.confirmPunch.category,'');
api.state.users=['田中'];
let n=0; const log=(type,time,category)=>({id:String(++n),name:'田中',type,time:'2026-09-17 '+time+':00',month:'2026-09',category,transport:'バス・電車（交通費あり）',memo:''});
api.state.logs=[log('出勤','09:00','アカデミー'),log('休憩開始','10:00','アカデミー'),log('休憩終了','10:15','アカデミー'),log('退勤','11:00','アカデミー'),log('出勤','11:00','ホームワイン'),log('退勤','12:00','ホームワイン'),log('出勤','13:00','アカデミー'),log('退勤','14:00','アカデミー'),log('出勤','14:00','その他（WT業務）'),log('退勤','15:00','その他（WT業務）')];
const st=api.calculateMonthlyStats('2026-09')['田中']; assert.equal(st.totalMinutes,285); assert.equal(st.categoryMinutes['アカデミー'],165); assert.equal(st.categoryMinutes['ホームワイン'],60); assert.equal(st.categoryMinutes['その他（WT業務）'],60); assert.equal(st.daysCount,1); assert.equal(st.paidTransportDays,1);
const rows=[['タイムスタンプ','名前','区分','打刻日時','対象月','システムID','交通機関','備考','業務区分','実働時間']];
const formulas=[];
const sheet={getLastRow:()=>rows.length,appendRow:r=>rows.push(r),getRange:(r,c,count=1,width=1)=>({getDisplayValues:()=>rows.slice(r-1,r-1+count).map(row=>row.slice(c-1,c-1+width)),setFormula(f){formulas.push({row:r,formula:f});return this},setNumberFormat(){return this}})};
const gas={LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},SpreadsheetApp:{flush(){}},console};vm.createContext(gas);vm.runInContext(fs.readFileSync(__dirname+'/Code.gs','utf8'),gas);
const p={name:'田中',type:'出勤',time:'2026-09-17 09:00:00',month:'2026-09',id:'a',category:'アカデミー',memo:'準備'};
assert.equal(gas.addLog_(sheet,p).ok,true); assert.equal(gas.addLog_(sheet,p).duplicate,true); assert.equal(gas.addLog_(sheet,{...p,id:'b'}).duplicate,true); assert.equal(rows.length,2);
assert.equal(gas.addLog_(sheet,{...p,id:'c',category:'ホームワイン'}).duplicate,undefined); assert.equal(rows.length,3);
assert.equal(gas.addLog_(sheet,{...p,name:'橋本'}).ok,false);
assert.equal(rows[1][8],'アカデミー');assert.equal(rows[1][7],'準備');assert.match(formulas[0].formula,/LOOKUP/);
async function checkRapidSubmit(){
  api.setRender(); let sends=0, resolveSend;
  api.setSend(()=>{sends++; return new Promise(resolve=>{resolveSend=resolve})});
  api.state.selectedUser='田中'; api.state.workCategory='アカデミー'; api.triggerPunchConfirmation('退勤'); el('checkoutMemo').value='講座準備';
  const pending=api.executePunch(); await api.executePunch(); assert.equal(sends,1); assert.equal(api.state.isPunchSubmitting,true); assert.equal(el('executePunch').disabled,true);
  resolveSend({ok:true,id:'request-1'}); await pending; assert.equal(api.state.isPunchSubmitting,false); assert.equal(api.state.logs[0].memo,'講座準備'); assert.equal(api.state.logs[0].category,'アカデミー');
}
checkRapidSubmit().then(()=>console.log('PASS: category selection, checkout memo, multi-session totals, breaks, one daily transport, idempotency, rapid clicks')).catch(e=>{console.error(e);process.exitCode=1});
