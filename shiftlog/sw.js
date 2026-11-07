// Bump CACHE_VERSION whenever any file in this folder changes.
const CACHE_VERSION = 'sjfd-shiftlog-v1.3.2';
const ASSETS = ['./', './index.html', './memo.js', './docx.js', './pdf-lib.min.js', './letterhead.js',
  './manifest.json', './icon.png', './apple-touch-icon.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE_VERSION).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(cached => {
    const net = fetch(e.request).then(res => { if (res.ok) caches.open(CACHE_VERSION).then(c => c.put(e.request, res.clone())); return res; }).catch(() => cached);
    return cached || net;
  }));
});
