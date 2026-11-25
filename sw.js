// Service worker: guarda a "casca" da app para abrir depressa e sem rede.
// Nunca guarda pedidos à Microsoft (login, OneDrive): esses vão sempre à rede.
const CACHE = 'financas-v1';
const SHELL = ['./', './index.html', './app.js', './xlsx-lite.js', './styles.css', './manifest.webmanifest',
  './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', './icons/favicon-32.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (url.origin !== location.origin) return; // Microsoft, OneDrive: sempre rede
  if (url.search.includes('code=')) return;   // regresso do login
  // App: rede primeiro (para receber atualizações), cache se estiver offline
  e.respondWith(
    fetch(e.request).then((r) => {
      const copy = r.clone();
      caches.open(CACHE).then((c) => c.put(e.request, copy));
      return r;
    }).catch(() => caches.match(e.request).then((r) => r || caches.match('./index.html')))
  );
});
