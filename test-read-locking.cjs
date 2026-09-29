// 読み取りで鍵を待たないこと、GET の打刻を拒むこと、書き込み・シート変更のあとにキャッシュを温めることを確かめる
const fs=require('fs'),vm=require('vm'),crypto=require('crypto'),assert=require('assert');
const cache=new Map(),props=new Map();let lockCalls=[],tryResult=true,opened=0;
const rows=[['2026-09-29 09:00:00','田中','出勤','2026-09-29 09:00:00','2026-09','id-1','','','','']];
const sheet={getLastRow:()=>rows.length+1,getName:()=>'勤怠マスタ',getRange:(r,c,n)=>({getDisplayValues:()=>rows.slice(r-2,r-2+n).map(x=>x.concat(Array(22-x.length).fill('')))})};
const ss={getSheetByName:n=>n==='勤怠マスタ'?sheet:null,getSheets:()=>[]};
const gas={Date,Set,JSON,
  CacheService:{getScriptCache:()=>({get:k=>cache.get(k)||null,put:(k,v)=>cache.set(k,v),remove:k=>cache.delete(k)})},
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v),deleteProperty:k=>props.delete(k)})},
  LockService:{getScriptLock:()=>({waitLock(ms){lockCalls.push(['wait',ms])},tryLock(ms){lockCalls.push(['try',ms]);return tryResult},releaseLock(){lockCalls.push(['release'])}})},
  Utilities:{getUuid:()=>crypto.randomUUID(),formatDate:()=>'2026-09-29 10:00:00'},
  SpreadsheetApp:{openById:()=>{opened++;return ss},flush(){}},
  ContentService:{MimeType:{JSON:'json',JAVASCRIPT:'js'},createTextOutput:c=>({setMimeType:()=>JSON.parse(c)})}};
vm.createContext(gas);vm.runInContext(fs.readFileSync(process.argv[2]||'Code.gs','utf8'),gas);
// 1) GET の打刻は拒否され、シートを開きもしない
opened=0;const r1=gas.doGet({parameter:{action:'add',name:'x',type:'出勤',time:'2026-09-29 09:00:00'}});
assert.strictEqual(r1.error,'post_required');assert.strictEqual(opened,0);
// 2) 読み取りは待たずに試す鍵（tryLock 0）で、取れたら索引を保存して鍵を外す
lockCalls=[];tryResult=true;const r2=gas.doGet({parameter:{action:'read',scope:'recent',fresh:'1'}});
assert.ok(r2.ok);assert.deepStrictEqual(lockCalls[0],['try',0]);assert.ok(lockCalls.some(c=>c[0]==='release'));
assert.ok(!lockCalls.some(c=>c[0]==='wait'),'読み取りで鍵を待ってはいけない');assert.ok(props.has('ATTENDANCE_READ_INDEX_meta'));
// 3) 鍵が取れなくても読めて、索引は保存せず、鍵も外さない
props.clear();lockCalls=[];tryResult=false;const r3=gas.doGet({parameter:{action:'read',scope:'recent',fresh:'1'}});
assert.ok(r3.ok);assert.strictEqual(r3.logs.length,1);assert.ok(!props.has('ATTENDANCE_READ_INDEX_meta'));assert.ok(!lockCalls.some(c=>c[0]==='release'));
// 鍵なしで読んだ結果はキャッシュに入れない（打刻直後に作り直したキャッシュを古い結果で上書きしないため）
cache.clear();tryResult=false;gas.doGet({parameter:{action:'read',scope:'recent',fresh:'1'}});assert.ok(!cache.has('attendance:recent'));
// 4) キャッシュは30分
assert.strictEqual(gas.RECENT_CACHE_SECONDS===undefined?vm.runInContext('RECENT_CACHE_SECONDS',gas):gas.RECENT_CACHE_SECONDS,1800);
// 5) シート変更のトリガーは鍵を待たず、キャッシュを作り直す
cache.clear();lockCalls=[];tryResult=true;gas.attendanceSheetChanged({});
assert.ok(!lockCalls.some(c=>c[0]==='wait'));assert.ok(lockCalls.some(c=>c[0]==='try'&&c[1]===10000),'索引を捨てる前に鍵を試す');assert.ok(cache.has('attendance:recent'),'トリガーのあとキャッシュが温まっている');
assert.ok(lockCalls.some(c=>c[0]==='try'&&c[1]===5000));
// 6) 索引は版（gen）ごとに保存し、切り替え前の途中の書き込みを読んでも古い版を最後まで一貫して読む
props.clear();
vm.runInContext("saveReadIndex_({version:1,mark:'A',pad:'x'.repeat(4500)})",gas);
const metaA=JSON.parse(props.get('ATTENDANCE_READ_INDEX_meta'));assert.ok(metaA.gen&&metaA.parts===3);
// 新しい版の部分だけ書かれ、meta はまだ古い版を指している状態
props.set('ATTENDANCE_READ_INDEX_zzzzzzzz_0','{"version":1,"mark":"B"');
assert.strictEqual(vm.runInContext('loadReadIndex_()',gas).mark,'A');
// 切り替え後は古い版の部分が消える
vm.runInContext("saveReadIndex_({version:1,mark:'C'})",gas);
assert.strictEqual(vm.runInContext('loadReadIndex_()',gas).mark,'C');
assert.ok(![...props.keys()].some(k=>k.startsWith('ATTENDANCE_READ_INDEX_'+metaA.gen)),'古い版の部分が残っていない');
// 古い版を読んでいる途中で消されたら、混ぜずに作り直し（null）になる
const metaC=JSON.parse(props.get('ATTENDANCE_READ_INDEX_meta'));props.delete('ATTENDANCE_READ_INDEX_'+metaC.gen+'_0');
assert.strictEqual(vm.runInContext('loadReadIndex_()',gas),null);
// 索引を捨てるときは部分も消す
props.delete('ATTENDANCE_READ_INDEX_zzzzzzzz_0'); // 上で手作業で入れた書きかけの版
vm.runInContext("saveReadIndex_({version:1,mark:'D'})",gas);vm.runInContext('invalidateAttendanceIndex_()',gas);
assert.ok(![...props.keys()].some(k=>k.startsWith('ATTENDANCE_READ_INDEX_')),'索引の部分がすべて消えている');
console.log('勤怠の変更点テスト: すべて合格');
