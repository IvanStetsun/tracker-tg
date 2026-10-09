// Keeps a copy of the app so it opens without internet (Android, desktop; iPhone's Telegram ignores service workers).
// Network first: when online you always get the latest version; the copy is used only when the network fails.
const CACHE = 'tracker-shell-v1';

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const ours = url.origin === self.location.origin;
  if (!(ours || url.hostname === 'telegram.org') || url.pathname.endsWith('version.txt')) return;

  event.respondWith(
    fetch(req)
      .then((res) => {
        if (res.ok || res.type === 'opaque') {
          const copy = res.clone();
          caches.open(CACHE).then((cache) => cache.put(req, copy)).catch(() => {});
        }
        return res;
      })
      .catch(() => caches.match(req, { ignoreSearch: true }).then((hit) => hit || Response.error()))
  );
});
