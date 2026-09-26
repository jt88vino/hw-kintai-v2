const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
// Home-screen guide: phones only, hidden once added or closed; Android uses its own install dialog.
const nodes=new Map();
function el(id){if(!nodes.has(id)){const n={id,value:'',textContent:'',innerHTML:'',classes:new Set(),classList:{add(v){n.classes.add(v)},remove(v){n.classes.delete(v)},toggle(v,on){if(on===undefined)on=!n.classes.has(v);on?n.classes.add(v):n.classes.delete(v);},contains(v){return n.classes.has(v)}},addEventListener(){},querySelectorAll:()=>[]};nodes.set(id,n);}return nodes.get(id);}
const store=new Map();let touch=true,standalone=false;
const client={Date,URLSearchParams,Set,Map,console,setTimeout:()=>1,clearTimeout(){},localStorage:{getItem:k=>store.has(k)?store.get(k):null,setItem:(k,v)=>store.set(k,v)},
  window:{addEventListener(){},crypto:require('node:crypto'),matchMedia:q=>({matches:q.includes('coarse')?touch:standalone})},navigator:{userAgent:'Mozilla/5.0 (iPhone; CPU iPhone OS 26_0 like Mac OS X) Safari/605.1.15',platform:'iPhone',maxTouchPoints:5},
  document:{getElementById:el,querySelectorAll:()=>[]}};
vm.createContext(client);
const script=fs.readFileSync('index.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
vm.runInContext(script.replace("if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();",`window.api={renderInstallTip,installStepsHtml,startInstall,setPrompt:p=>{installPrompt=p;}};`),client);
const api=client.window.api;
(async()=>{
 api.renderInstallTip();assert(!el('installTip').classes.has('hidden'),'a phone in the browser sees the guide');assert.equal(el('installBtn').textContent,'追加のしかたを見る');
 assert.match(api.installStepsHtml(),/iPhone・iPad/);assert.match(api.installStepsHtml(),/ホーム画面に追加/);assert.doesNotMatch(api.installStepsHtml(),/Android/);
 await api.startInstall();assert(el('installModal').classes.has('show'),'without a native prompt the steps open');
 standalone=true;api.renderInstallTip();assert(el('installTip').classes.has('hidden'),'opened from the home screen: no guide');standalone=false;
 touch=false;api.renderInstallTip();assert(el('installTip').classes.has('hidden'),'desktop: no guide');touch=true;
 client.navigator.userAgent='Mozilla/5.0 (Linux; Android 15) Chrome/140 Mobile';assert.match(api.installStepsHtml(),/Android/);assert.doesNotMatch(api.installStepsHtml(),/iPhone・iPad/);
 let prompted=0;api.setPrompt({prompt:async()=>{prompted++;},userChoice:Promise.resolve({outcome:'accepted'})});api.renderInstallTip();assert.equal(el('installBtn').textContent,'ホーム画面に追加する');
 await api.startInstall();assert.equal(prompted,1);assert.equal(store.get('attendance_install_tip'),'"done"');api.renderInstallTip();assert(el('installTip').classes.has('hidden'),'added: the guide stays away');
 const man=JSON.parse(fs.readFileSync('manifest.webmanifest','utf8'));assert.equal(man.display,'standalone');assert.equal(man.short_name,'HW勤怠');
 for(const i of man.icons) assert(fs.existsSync('.'+i.src),i.src);assert(fs.existsSync('icons/apple-touch-icon.png'));
 const head=fs.readFileSync('index.html','utf8');assert.match(head,/rel="manifest" href="\/manifest.webmanifest"/);assert.match(head,/rel="apple-touch-icon"/);
 console.log('test-install: OK');
})().catch(e=>{console.error(e);process.exit(1);});
