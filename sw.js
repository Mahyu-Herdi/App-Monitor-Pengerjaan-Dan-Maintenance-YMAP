// ================================================================
//  sw.js — Service Worker SMM Dapur MBG
//  Ganti versi CACHE_NAME setiap kali update file untuk paksa refresh
// ================================================================
const CACHE_NAME = 'smm-dapur-v2';
const ASSETS = ['./', './index.html', './config.js', './manifest.json',
                './sp.html', './invoice.html', './lampiran.html'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME)
      .then(c => c.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Lewati semua POST (ke GAS) dan request ke script.google.com
  if (e.request.method !== 'GET') return;
  if (e.request.url.includes('script.google.com')) return;

  e.respondWith(
    caches.match(e.request).then(cached => {
      const network = fetch(e.request).then(res => {
        if (res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => {});
      return cached || network;
    })
  );
});
