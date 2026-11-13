const V='finanzas-v9';
const ARCH=['config.js','jspdf.umd.min.js','jspdf.plugin.autotable.min.js','./','index.html','worker.js','shim.js','gs-bundle.js','prelude.js','chart.umd.min.js','lucide.min.js','manifest.webmanifest','icon-192.png','icon-512.png','icon-maskable-512.png','apple-touch-icon.png','codigo-hoja.gs','codigo-correos.gs'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>Promise.all(ARCH.map(u=>fetch(new Request(u,{cache:'reload'})).then(r=>{if(!r.ok)throw new Error(u);return c.put(u,r)})))).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==V).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
/* Con internet siempre se usa la versión más nueva (así index, motor y código nunca quedan de versiones distintas); sin internet, la guardada.
   Se guarda sin el ?v=… para que la caché no crezca con cada versión. */
self.addEventListener('fetch',e=>{
  const r=e.request,u=new URL(r.url);
  if(r.method!=='GET'||u.origin!==location.origin)return;
  const clave=u.origin+u.pathname;
  e.respondWith(caches.open(V).then(c=>{
    const red=fetch(r,{cache:'no-cache'}).then(x=>{if(x&&x.ok&&x.type==='basic')c.put(clave,x.clone());return x});
    const limite=new Promise((_,no)=>setTimeout(()=>no(new Error('lento')),5000));
    return Promise.race([red,limite]).catch(()=>c.match(clave,{ignoreSearch:true}).then(hit=>hit||(r.mode==='navigate'?c.match('index.html'):null)).then(hit=>hit||red))
  }))
});
