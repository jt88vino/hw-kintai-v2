const fs=require('fs'),vm=require('vm'),assert=require('assert/strict'),crypto=require('crypto');
const props=new Map(),reads=[];
const rows=[Array(21).fill('header')];
function row(name,type,time,id,category=''){const r=Array(21).fill('');Object.assign(r,{1:name,2:type,3:time,4:time.slice(0,7),5:id,8:category});return r;}
rows.push(row('長期休み','退勤','2020-01-01 18:00:00','old'));
rows.push(row('日跨ぎ','出勤','2020-01-31 23:00:00','open','業務配分'));
for(let year=2020;year<2027;year++)for(let month=1;month<=12;month++)for(let day=1;day<=20;day++)rows.push(row('通常','退勤',`${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')} 18:00:00`,`${year}-${month}-${day}`));
const sheet={getLastRow:()=>rows.length,getRange:(start,col,count,width)=>({getDisplayValues:()=>{reads.push({start,count});return rows.slice(start-1,start-1+count).map(r=>r.slice(col-1,col-1+width));}})};
const ctx={Date,Set,PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v),deleteProperty:k=>props.delete(k)})},Utilities:{getUuid:()=>crypto.randomUUID(),formatDate:()=> '2026-12-31'},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})}};
vm.createContext(ctx);vm.runInContext(fs.readFileSync('Code.gs','utf8'),ctx);ctx.readRoster_=()=>({users:['長期休み','日跨ぎ','通常'],transportationCosts:{}});
assert.equal(ctx.indexTime_('2026/9/17 9:02:03'),'2026-09-17 09:02:03');
let data=ctx.readData_({},sheet,{scope:'recent'});assert.equal(data.schemaVersion,4);assert.equal(data.logs.length,52);assert(data.logs.some(l=>l.id==='old'));assert(data.logs.some(l=>l.id==='open'));assert.equal(data.months.length,84);
reads.length=0;data=ctx.readData_({},sheet,{scope:'recent'});assert.equal(reads.reduce((n,r)=>n+r.count,0),52,'warm initial read bounded to recent plus status/session');
reads.length=0;data=ctx.readData_({},sheet,{scope:'month',month:'2023-06'});assert.equal(data.logs.filter(l=>l.month==='2023-06').length,20);assert(reads.reduce((n,r)=>n+r.count,0)<30,'monthly lookup must not scan whole history');
assert.equal(ctx.readData_({},sheet,{scope:'month',month:'2023-99'}).ok,false);
rows.push(row('通常','出勤','2027-01-01 09:00:00','new'));reads.length=0;data=ctx.readData_({},sheet,{scope:'recent'});assert(data.logs.some(l=>l.id==='new'));assert(reads.reduce((n,r)=>n+r.count,0)<60,'append indexed from new rows only');
// Direct edit trigger and app deletion invalidate row references, including older statuses.
rows[1][2]='出勤';ctx.attendanceSheetChanged({});data=ctx.readData_({},sheet,{scope:'recent'});assert.equal(data.logs.find(l=>l.id==='old').type,'出勤');
rows.splice(1,1);ctx.invalidateAttendanceIndex_();data=ctx.readData_({},sheet,{scope:'recent'});assert(!data.logs.some(l=>l.id==='old'));assert(data.logs.some(l=>l.id==='open'));
rows.push(row('過去追記','退勤','2021-02-02 12:00:00','backdated'));data=ctx.readData_({},sheet,{scope:'month',month:'2021-02'});assert(data.logs.some(l=>l.id==='backdated'));assert.equal(data.logs.filter(l=>l.month==='2021-02').length,21);
// UI receives independently scoped data; switching months cannot overwrite current status.
const source=fs.readFileSync('index.html','utf8').match(/<script>([\s\S]*)<\/script>/)[1];
const els=new Map();const el=id=>{if(!els.has(id))els.set(id,{innerHTML:'',value:'',classList:{add(){},remove(){},toggle(){}},querySelector:()=>null,querySelectorAll:()=>[]});return els.get(id)};
const client={console,URLSearchParams,Date,setTimeout,clearTimeout,localStorage:{getItem:()=>null,setItem(){}},window:{addEventListener(){}},document:{getElementById:el,querySelectorAll:()=>[]}};
vm.createContext(client);vm.runInContext(source.replace("if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();",`window.test={state,applyGasReadData,loadSelectedMonth,calculateMonthlyStats,setRead:fn=>readGASDataOnce=fn,prepare:()=>renderDashboard=()=>{}};`),client);
const api=client.window.test;api.prepare();api.state.selectedMonth='2026-09';api.applyGasReadData({ok:true,scope:'recent',logs:[{id:'now',name:'通常',type:'出勤',time:'2026-09-17 09:00:00',month:'2026-09'}],months:['2026-09','2020-01']});
api.applyGasReadData({ok:true,scope:'month',month:'2026-09',logs:[{id:'historic',name:'通常',type:'退勤',time:'2026-09-01 18:00:00',month:'2026-09'}]});assert.equal(api.state.logs[0].id,'now');assert.equal(api.state.monthLogs[0].id,'historic');
(async()=>{let resolves=[];api.setRead(()=>new Promise(r=>resolves.push(r)));const a=api.loadSelectedMonth();api.state.selectedMonth='2020-01';const b=api.loadSelectedMonth();resolves[1]({ok:true,scope:'month',month:'2020-01',logs:[]});await b;resolves[0]({ok:true,scope:'month',month:'2026-09',logs:[]});await a;assert.equal(api.state.loadedMonth,'2020-01');assert.equal(api.state.logs[0].id,'now');console.log('PASS: 1683-row history, bounded 52-row warm read, targeted month, append indexing, edit/delete invalidation, backdated month, isolated scopes, out-of-order responses');})().catch(e=>{console.error(e);process.exitCode=1});
