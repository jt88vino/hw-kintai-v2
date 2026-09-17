const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),crypto=require('node:crypto');
const cache=new Map(); let configured=true, writes=0;
const ctx={PropertiesService:{getScriptProperties:()=>({getProperty:()=>configured?crypto.createHash('sha256').update('test-only-password').digest('hex'):null})},CacheService:{getScriptCache:()=>({get:k=>cache.get(k),put:(k,v)=>cache.set(k,v),remove:k=>cache.delete(k)})},Utilities:{DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},computeDigest:(_,v)=>[...crypto.createHash('sha256').update(v).digest()],getUuid:()=>crypto.randomUUID()},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},SpreadsheetApp:{openById(){writes++;throw Error('No production writes in test')}},ContentService:{MimeType:{JSON:'json',JAVASCRIPT:'js'},createTextOutput:content=>({setMimeType:()=>JSON.parse(content)})}};
vm.createContext(ctx);vm.runInContext(fs.readFileSync('Code.gs','utf8'),ctx);
configured=false;assert.equal(ctx.loginAdmin_({adminPassword:'test-only-password'}).error,'admin_not_configured');configured=true;
assert.equal(ctx.loginAdmin_({adminPassword:'incorrect'}).ok,false);
const login=ctx.loginAdmin_({adminPassword:'test-only-password'});assert.equal(login.ok,true);assert.equal(ctx.validateAdmin_({adminToken:login.token},'delete'),null);
for(const action of ['delete','addUser','deleteUser']){
 assert.equal(ctx.doGet({parameter:{action}}).error,'post_required');
 const receipt=crypto.randomUUID();const out=ctx.doPost({parameter:{action,receipt}});assert.equal(out.error,'unauthorized');assert.equal(ctx.doGet({parameter:{action:'adminResult',receipt}}).error,'unauthorized');
}
assert.equal(writes,0);assert.equal(ctx.deleteLog_({},{}).error,'unauthorized');
cache.delete('admin:'+login.token);assert.equal(ctx.validateAdmin_({adminToken:login.token},'delete').error,'unauthorized');
for(let i=0;i<20;i++)ctx.loginAdmin_({adminPassword:'incorrect'});assert.equal(ctx.loginAdmin_({adminPassword:'test-only-password'}).error,'rate_limited');
const html=fs.readFileSync('index.html','utf8');assert.doesNotMatch(html,/const ADMIN_PASSWORD\s*=/);assert.doesNotMatch(html,/adminPassword:ADMIN_PASSWORD/);
console.log('PASS: server-only authentication, missing config fails closed, wrong password, token expiry, POST-only admin actions, denied writes, receipt polling, throttling');
