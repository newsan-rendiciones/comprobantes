/* Service worker mínimo:
   - habilita "Instalar" en Android/Chrome
   - cachea SOLO el envoltorio (index/manifest/íconos)
   - NUNCA cachea la app de Google: esa va siempre en vivo desde el iframe */

const CACHE = 'comprobantes-v1';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', function (e) {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE).then(function (c) {
      return c.addAll(SHELL).catch(function () {}); // si falta un ícono, no rompe
    })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) { if (k !== CACHE) return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function (e) {
  var url = e.request.url;
  // Dejar pasar en vivo todo lo de Google (contenido dinámico del iframe)
  if (url.indexOf('script.google.com') !== -1 || url.indexOf('googleusercontent.com') !== -1) return;
  e.respondWith(
    caches.match(e.request).then(function (hit) { return hit || fetch(e.request); })
  );
});
