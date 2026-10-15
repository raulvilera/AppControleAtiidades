// Guarda a casca do app para abrir rápido e offline; os dados vêm sempre da planilha.
var V='atividades-v9', ARQ=['./','index.html','manifest.webmanifest','icon-192.png','icon-512.png'];
self.addEventListener('install',function(e){ e.waitUntil(caches.open(V).then(function(c){return c.addAll(ARQ);}).then(function(){return self.skipWaiting();})); });
self.addEventListener('activate',function(e){ e.waitUntil(caches.keys().then(function(ks){return Promise.all(ks.filter(function(k){return k!==V;}).map(function(k){return caches.delete(k);}));}).then(function(){return self.clients.claim();})); });
self.addEventListener('fetch',function(e){
  var r=e.request; if(r.method!=='GET'||new URL(r.url).origin!==location.origin) return;
  e.respondWith(fetch(r).then(function(res){ var cp=res.clone(); caches.open(V).then(function(c){c.put(r,cp);}); return res; }).catch(function(){ return caches.match(r).then(function(m){ return m||caches.match('index.html'); }); }));
});
