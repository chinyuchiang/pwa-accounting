const CACHE_NAME = 'ledger-v1';
const PRECACHE = ['/', '/index.html', '/manifest.json', '/icon.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(PRECACHE)));
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
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  if (url.hostname.includes('script.google.com')) return;
  if (url.hostname.includes('cdn.jsdelivr.net')) return;

  e.respondWith(
    caches.match(e.request).then(cached => {
      const network = fetch(e.request).then(res => {
        if (res.ok && url.origin === location.origin) {
          caches.open(CACHE_NAME).then(c => c.put(e.request, res.clone()));
        }
        return res;
      });
      return cached || network;
    })
  );
});
