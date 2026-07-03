var CACHE = 'po-manager-v49';
var SHELL = ['/po-manager/', '/po-manager/index.html', '/po-manager/manifest.json', '/po-manager/icon-192.png', '/po-manager/icon-512.png', '/po-manager/apple-touch-icon.png'];

self.addEventListener('install', function(e) {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(function(c) { return c.addAll(SHELL); }));
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      var old = keys.filter(function(k) { return k !== CACHE; });
      var isUpdate = old.length > 0;
      return Promise.all(old.map(function(k) { return caches.delete(k); }))
        .then(function() { return clients.claim(); })
        .then(function() {
          if (isUpdate) {
            return self.clients.matchAll({ type: 'window' }).then(function(cls) {
              cls.forEach(function