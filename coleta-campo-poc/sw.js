const CACHE_NAME = 'coleta-campo-v1';
const ASSETS = ['./', './index.html', './app.js', './manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return; // POST de sync vai direto pra rede
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).catch(() => caches.match('./index.html')))
  );
});

// Background Sync (bônus - só Chrome/Android; iOS ignora isso silenciosamente)
self.addEventListener('sync', e => {
  if (e.tag === 'sync-registros') {
    e.waitUntil(
      self.clients.matchAll().then(clients => clients.forEach(c => c.postMessage({ tipo: 'TRIGGER_SYNC' })))
    );
  }
});