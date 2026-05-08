var CACHE = 'po-manager-v1';
var SHELL = ['/po-manager/', '/po-manager/index.html', '/po-manager/manifest.json', '/po-manager/icon-192.png', '/po-manager/icon-512.png', '/po-manager/apple-touch-icon.png'];

self.addEventListener('install', function(e) {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(function(c) { return c.addAll(SHELL); }));
});

self.addEventListener('activate', function(e) {
  e.waitUntil(clients.claim());
});

// Fetch handler — required for Chrome PWA installability
self.addEventListener('fetch', function(e) {
  // Only intercept same-origin requests (GitHub Pages shell)
  // Let all Google Script requests pass through normally
  if (e.request.url.startsWith(self.location.origin)) {
    e.respondWith(
      caches.match(e.request).then(function(cached) {
        return cached || fetch(e.request);
      })
    );
  }
});
