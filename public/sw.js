const CACHE_NAME = 'art1s-store-v3';
const urlsToCache = [
  '/',
  '/index.html',
  '/manifest.json',
  '/art1s-bg.svg',
  '/art1s-outline.svg'
];

self.addEventListener('install', (e) => {
  self.skipWaiting(); // Force active new service worker immediately
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(urlsToCache))
  );
});

self.addEventListener('activate', (e) => {
  // Clean up old caches
  e.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cache) => {
          if (cache !== CACHE_NAME) {
            return caches.delete(cache);
          }
        })
      );
    }).then(() => self.clients.claim()) // Force service worker to take control of page
  );
});

self.addEventListener('fetch', (e) => {
  // Only handle GET requests (bypass POST, PUT, DELETE, etc.)
  if (e.request.method !== 'GET') {
    return;
  }

  // Bypass cache for Google Apps Script API calls
  if (e.request.url.includes('script.google.com')) {
    e.respondWith(fetch(e.request));
    return;
  }

  // Network-First Strategy: Fetch from internet first, save/update cache on success
  e.respondWith(
    fetch(e.request)
      .then((response) => {
        // If successful, cache it
        if (response && response.status === 200 && (response.type === 'basic' || response.type === 'cors')) {
          const responseToCache = response.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(e.request, responseToCache);
          });
        }
        return response;
      })
      .catch(() => {
        // If offline/error, serve from cache
        return caches.match(e.request);
      })
  );
});
