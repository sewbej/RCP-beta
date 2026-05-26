const CACHE_NAME = 'skaner-v4-ultra';
const ASSETS = [
  'index.html',
  'html5-qrcode.min.js',
  'xlsx.full.min.js',
  'manifest.json'
];

// Instalacja - natychmiastowe zapisanie
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS);
    }).then(() => self.skipWaiting())
  );
});

// Aktywacja - przejęcie kontroli nad aplikacją
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cache) => {
          if (cache !== CACHE_NAME) return caches.delete(cache);
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Obsługa zapytań - TOTALNY OFFLINE
self.addEventListener('fetch', (event) => {
  event.respondWith(
    caches.match(event.request, { ignoreSearch: true }).then((response) => {
      // Zwróć z pamięci, jeśli masz. Jeśli nie, spróbuj pobrać.
      return response || fetch(event.request).catch(() => {
        // Jeśli nie ma w pamięci i nie ma sieci, a to strona - zwróć index.html
        if (event.request.mode === 'navigate') {
          return caches.match('index.html');
        }
      });
    })
  );
});
