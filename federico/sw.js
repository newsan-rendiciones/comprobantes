/* Service worker de la app de Federico (scope: /federico/) */
const CACHE = 'comprobantes-federico-v1';
const SHELL = ['./', './index.html', './manifest.json', '../icon-192.png', '../icon-512.png'];
self.addEventListener('install', function (e) {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL).catch(function () {}); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.map(function (k) { if (k !== CACHE) return caches.delete(k); }));
  }));
  self.clients.claim();
});
self.addEventListener('fetch', function (e) {
  var url = e.request.url;
  if (url.indexOf('script.google.com') !== -1 || url.indexOf('googleusercontent.com') !== -1) return;
  e.respondWith(caches.match(e.request).then(function (hit) { return hit || fetch(e.request); }));
});
