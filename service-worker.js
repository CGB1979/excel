// service-worker.js
const CACHE_NAME = "excel-main-v1.2.3.2"; // actualizá aquí a la versión nueva
const APP_SHELL = [
  "./",
  "index.html",
  "css/estilos.css",
  "js/app.js",
  "js/version.js",
  "manifest.json"
];

// Instalación: precacheo
self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

// Activación: limpia caches viejos
self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

// Fetch: estrategia cache-first con actualización en background
self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;

  event.respondWith(
    caches.match(event.request).then(cached => {
      const networkFetch = fetch(event.request)
        .then(response => {
          // Si la respuesta es válida, actualizamos cache
          if (response && response.status === 200) {
            caches.open(CACHE_NAME).then(cache => {
              // Evitamos cachear cosas cross-origin problemáticas si fuese necesario
              try { cache.put(event.request, response.clone()); } catch (e) { /* ignore */ }
            });
          }
          return response;
        })
        .catch(() => cached); // si falla la red, devolvemos cached (si existe)

      // Devolver cached si existe inmediatamente, sino la promesa de red
      return cached || networkFetch;
    })
  );
});

// Permitir que la página solicite activar el SW inmediatamente
self.addEventListener("message", event => {
  if (event.data && event.data.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
});
