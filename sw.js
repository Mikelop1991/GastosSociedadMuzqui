// Sociedad Muzki: guarda la app para que abra sin conexión. Los datos van aparte (localStorage).
const CACHE = 'muzki-v7';
const BASE = ['./', 'index.html', 'manifest.json', 'plantilla.xlsx', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(BASE))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.hostname.includes('script.google')) return; // los datos siempre a la red
  // red primero (para recibir mejoras), caché si no hay conexión
  e.respondWith(fetch(e.request).then(r => {
    if (r.ok && (u.origin === location.origin || u.hostname.includes('fonts') || u.hostname.includes('cdnjs'))) {
      const copia = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copia));
    }
    return r;
  }).catch(() => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match('index.html'))));
});
