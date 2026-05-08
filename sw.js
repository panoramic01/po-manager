// Minimal service worker — required for Chrome to enable standalone PWA mode
self.addEventListener('install', function(e) {
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(clients.claim());
});

// Pass all requests straight through to the network
self.addEventListener('fetch', function(e) {
  e.respondWith(fetch(e.request));
});
