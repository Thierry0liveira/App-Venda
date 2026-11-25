// Service worker: faz o app funcionar offline (cache dos arquivos).
// Aumente V a cada nova versão do app para renovar o cache.
const V='gc-v4',F=['./','index.html','styles.css','app.js','xlsx.min.js','icon.svg','manifest.webmanifest'];
// Instalação: guarda os arquivos do app no cache para abrir sem internet.
self.addEventListener('install',e=>e.waitUntil(caches.open(V).then(c=>c.addAll(F)).then(()=>self.skipWaiting())));
// Ativação: apaga caches de versões antigas.
self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==V).map(x=>caches.delete(x)))).then(()=>clients.claim())));
// Requisições: responde do cache primeiro; se não houver, busca na rede; offline, cai no index.html.
self.addEventListener('fetch',e=>{if(e.request.method!=='GET')return;
e.respondWith(caches.match(e.request,{ignoreSearch:true}).then(r=>r||fetch(e.request).catch(()=>caches.match('index.html'))))});
