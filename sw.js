const V='finanzas-v5';
const ARCH=['config.js','jspdf.umd.min.js','jspdf.plugin.autotable.min.js','./','index.html','worker.js','shim.js','gs-bundle.js','prelude.js','chart.umd.min.js','lucide.min.js','manifest.webmanifest','icon-192.png','icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>c.addAll(ARCH)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==V).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const r=e.request;
  if(r.method!=='GET'||new URL(r.url).origin!==location.origin)return;
  const u=new URL(r.url);
  if(/(\/|index\.html|config\.js)$/.test(u.pathname)){
    e.respondWith(caches.open(V).then(c=>fetch(r).then(x=>{if(x&&x.ok)c.put(r,x.clone());return x}).catch(()=>c.match(r,{ignoreSearch:true}))));return
  }
  e.respondWith(caches.open(V).then(c=>c.match(r,{ignoreSearch:true}).then(hit=>{
    const red=fetch(r).then(x=>{if(x&&x.ok)c.put(r,x.clone());return x}).catch(()=>hit);
    return hit||red;
  })));
});
