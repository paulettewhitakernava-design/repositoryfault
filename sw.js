// Service Worker de Kashly: guarda en caché el "cascarón" de la app (HTML,
// íconos) para que abra sin internet. Los datos (movimientos, config) NUNCA
// se cachean aquí — eso lo maneja index.html con localStorage, porque
// necesita lógica propia para encolar y reintentar cambios pendientes.
const CACHE_NAME = 'kashly-shell-v1';
const ASSETS = [
  './',
  './index.html',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_NAME; }).map(function(k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(event) {
  const url = new URL(event.request.url);

  // La API (Apps Script) nunca se sirve desde caché: siempre debe intentar
  // red primero para no mostrar datos viejos como si fueran nuevos.
  if (url.hostname.indexOf('script.google.com') !== -1) return;
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request).then(function(cached) {
      const fetchPromise = fetch(event.request).then(function(networkResponse) {
        if (networkResponse && networkResponse.ok) {
          caches.open(CACHE_NAME).then(function(cache) { cache.put(event.request, networkResponse.clone()); });
        }
        return networkResponse;
      }).catch(function() { return cached; });
      return cached || fetchPromise;
    })
  );
});
