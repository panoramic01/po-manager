// Minimal service worker — just enough for Chrome to enable standalone PWA mode
self.addEventListener('install', function() { self.skipWaiting(); });
self.addEventListener('activate', function(e) { e.waitUntil(clients.claim()); });
