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
 console.log('PASS: legacy API: stable id, errors fail closed, credentials excluded, fixed upstream, private no-store');
})().catch(e=>{console.error(e);process.exitCode=1});
