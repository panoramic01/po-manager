var CACHE = 'po-manager-v199';
var SHELL = ['/', '/index.html', '/manifest.json', '/icon-192.png?v=126', '/icon-512.png?v=126', '/apple-touch-icon.png?v=126', '/panoramic-logo.png', '/panoramic-roofline.png'];

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
              cls.forEach(function(c) { c.postMessage({ type: 'SW_UPDATED' }); });
            });
          }
        });
    })
  );
});

self.addEventListener('push', function(e) {
  var data = {};
  try { data = e.data ? e.data.json() : {}; } catch (err) {}
  var payload = data.data || data; // FCM v1 data-only payload is nested under "data"
  var title = payload.title || 'Panoramic Ops';
  var body  = payload.body  || '';
  var url   = payload.url   || './';
  e.waitUntil(self.registration.showNotification(title, {
    body: body,
    icon: '/icon-192.png?v=126',
    badge: '/icon-192.png?v=126',
    data: { url: url }
  }));
});

self.addEventListener('notificationclick', function(e) {
  e.notification.close();
  var url = (e.notification.data && e.notification.data.url) || './';
  e.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function(cls) {
      for (var i = 0; i < cls.length; i++) {
        if ('focus' in cls[i]) { cls[i].navigate(url); return cls[i].focus(); }
      }
      return clients.openWindow(url);
    })
  );
});

self.addEventListener('fetch', function(e) {
  // Network-first for navigation (HTML) — always load the latest app on open
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request).catch(function() {
        return caches.match(e.request);
      })
    );
    return;
  }
  // Cache-first for all other assets (icons, manifest, etc.)
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      return cached || fetch(e.request);
    })
  );
});
