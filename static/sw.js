self.addEventListener('install', (e) => {
  console.log('AURAS Service Worker Installed');
});
self.addEventListener('fetch', (e) => {
  e.respondWith(fetch(e.request));
});