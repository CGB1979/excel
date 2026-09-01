const CACHE_NAME="convertidor-excel-v1.0.0.h";
const ASSETS=["./","./index.html","./css/estilos.css","./js/version.js","./js/configuracionColumnas.js","./js/modal.js","./js/excel.js","./js/unificador.js","./js/codigosBarras.js","./js/app.js","./manifest.json"];
self.addEventListener("install",e=>e.waitUntil(caches.open(CACHE_NAME).then(c=>c.addAll(ASSETS)).then(()=>self.skipWaiting())));
self.addEventListener("activate",e=>e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE_NAME).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener("fetch",e=>{if(e.request.method!=="GET")return;e.respondWith(fetch(e.request).then(r=>{const c=r.clone();caches.open(CACHE_NAME).then(x=>x.put(e.request,c));return r}).catch(()=>caches.match(e.request)))});
