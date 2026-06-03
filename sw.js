const CACHE_NAME = 'smm-dapur-cache-v2';
const ASSETS = [
  './',
  './index.html',
  './config.js',
  './manifest.json',
  './invoice.html',
  './sp.html',
  './lampiran.html'
];

// Install Service Worker & Cache Assets Offline
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      console.log('Caching App Shell...');
      return cache.addAll(ASSETS);
    }).then(() => self.skipWaiting())
  );
});

// Aktivasi & Pembersihan Cache Lama
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.map(key => {
          if (key !== CACHE_NAME) {
            console.log('Menghapus Cache Lama:', key);
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Strategi Fetch: Cache First, Fallback to Network
self.addEventListener('fetch', e => {
  // Jangan intercept request POST API ke Google Sheets
  if (e.request.method === 'POST') return;

  e.respondWith(
    caches.match(e.request).then(cachedResponse => {
      if (cachedResponse) {
        fetch(e.request).then(networkResponse => {
          if (networkResponse.status === 200) {
            caches.open(CACHE_NAME).then(cache => cache.put(e.request, networkResponse));
          }
        }).catch(() => {});
        return cachedResponse;
      }
      return fetch(e.request);
    })
  );
});
