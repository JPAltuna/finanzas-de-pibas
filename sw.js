/* Finanzas de pibas — service worker: cachea la app para que abra sin internet */
var CACHE = 'fdp-v7';
var CORE = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(CORE); }));
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function (e) {
  if (e.request.method !== 'GET') return;
  var url = new URL(e.request.url);
  // Las llamadas a Apps Script van siempre a la red, nunca al caché.
  if (url.hostname.indexOf('script.google') >= 0) return;
  e.respondWith(
    fetch(e.request).then(function (resp) {
      if (url.origin === location.origin && resp.ok) {
        var copia = resp.clone();
        caches.open(CACHE).then(function (c) { c.put(e.request, copia); });
      }
      return resp;
    }).catch(function () {
      return caches.match(e.request, { ignoreSearch: true });
    })
  );
});
