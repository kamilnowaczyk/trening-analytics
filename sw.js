/* Service Worker — Trening Analytics
   Cache „app shell" → aplikacja działa offline (np. słabe wifi na siłowni).
   Pliki danych użytkownika NIE są cache'owane (są wczytywane lokalnie). */
const CACHE = 'trening-analytics-v1';
const ASSETS = [
  './',
  './index.html',
  './app/styles.css',
  './app/app.js',
  './lib/jszip.min.js',
  './lib/chart.umd.min.js',
  './lib/xlsx-engine.js',
  './lib/training-model.js',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  // tylko zasoby z tego samego origin (nie czcionki Google itp.)
  if (new URL(req.url).origin !== location.origin) return;
  e.respondWith(
    caches.match(req).then((hit) => hit || fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
      return res;
    }).catch(() => caches.match('./index.html')))
  );
});
