const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const handler=require('./api/attendance');
function response(){return {headers:{},code:200,setHeader(k,v){this.headers[k]=v;},status(n){this.code=n;return this;},json(data){this.data=data;return this;}};}
let calls=[];global.fetch=async(url,options)=>{calls.push({url,options});return {ok:true,json:async()=>({ok:true,action:'add',id:'stable-id',memo:'保存済み',stack:'do-not-relay'})};};
const post={method:'POST',headers:{'content-type':'application/json',origin:'https://hw-kintai-v2.vercel.app'},body:{action:'add',requestId:'stable-id',name:'テスト専用',type:'出勤',time:'2026-09-18 09:00:00',month:'2026-09',allocations:'[]'}};
(async()=>{
 let res=response();await handler(post,res);assert.equal(res.data.ok,true);assert.equal(calls.length,1);assert.equal(calls[0].options.method,'POST');assert.equal(calls[0].options.body.get('requestId'),'stable-id');assert(!calls[0].url.includes('テスト専用'));assert(!('stack' in res.data));assert.equal(res.headers['Cache-Control'],'private, no-store');
 res=response();await handler({...post,body:{...post.body,action:'adminLogin',adminPassword:'test-only'}},res);assert.equal(res.code,400);assert.equal(calls.length,1,'credentials not forwarded');
 res=response();await handler({...post,headers:{...post.headers,origin:'https://unrelated.example'}},res);assert.equal(res.code,403);
 res=response();await handler({method:'GET',query:{action:'delete'},headers:{}},res);assert.equal(res.code,400);
 res=response();await handler({method:'GET',query:{scope:'month',month:'2026-99'},headers:{}},res);assert.equal(res.code,400);
 res=response();await handler({method:'GET',query:{scope:'recent',url:'https://unrelated.example'},headers:{}},res);assert.equal(calls.length,2);assert(calls[1].url.startsWith('https://script.google.com/'));assert(!calls[1].url.includes('unrelated'));
 global.fetch=async()=>{const error=new Error('sensitive detail');error.name='TimeoutError';throw error};res=response();await handler(post,res);assert.equal(res.code,504);assert.equal(res.data.ok,false);assert(!JSON.stringify(res.data).includes('sensitive detail'));
 // Browser: one successful POST, no receipt GET and no automatic replay on timeout.
 const nodes=new Map();const el=id=>{if(!nodes.has(id))nodes.set(id,{textContent:'',innerHTML:'',classList:{toggle(){},add(){},remove(){}},querySelectorAll:()=>[]});return nodes.get(id)};
 const client={Date,URL,URLSearchParams,AbortController,Object,console,setTimeout,clearTimeout,localStorage:{getItem:()=>null,setItem(){}},window:{location:{hostname:'hw-kintai-v2.vercel.app'},addEventListener(){},crypto:{randomUUID:()=> 'receipt'}},document:{getElementById:el,querySelectorAll:()=>[]}};
 let browserCalls=[];client.fetch=async(url,options)=>{browserCalls.push({url,options});return {ok:true,json:async()=>({ok:true,id:'stable-id'})}};
 vm.createContext(client);const source=fs.readFileSync('index.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];vm.runInContext(source.replace("if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();",'window.test={state,adminRequest,performSync,prepare:()=>{renderAll=()=>{};setSyncing=()=>{};},setRead:fn=>readGASDataOnce=fn};'),client);
 const api=client.window.test;await api.adminRequest(post.body);assert.equal(browserCalls.length,1);assert.equal(browserCalls[0].url,'/api/attendance');assert.equal(browserCalls[0].options.method,'POST');
 client.fetch=async()=>{browserCalls.push('failed');throw new Error('timeout')};await assert.rejects(()=>api.adminRequest(post.body),/timeout/);assert.equal(browserCalls.length,2,'uncertain write never automatically resubmitted');
 api.prepare();let finish;api.setRead(()=>new Promise(r=>finish=r));const sync=api.performSync('auto');api.state.mutationVersion++;api.state.logs=[{id:'confirmed'}];finish({ok:true,scope:'recent',logs:[]});await sync;assert.equal(api.state.logs[0].id,'confirmed','old response cannot erase confirmed punch');
 console.log('PASS: one POST response, stable id, no replay on timeout, errors fail closed, credentials excluded, fixed upstream, private no-store, stale read cannot undo confirmed punch');
})().catch(e=>{console.error(e);process.exitCode=1});
