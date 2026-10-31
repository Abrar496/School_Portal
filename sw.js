// Keeps the page itself available offline; the notices come from the network (the page keeps its own last copy).
const CACHE = "board-v2";
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(["./", "index.html", "manifest.webmanifest", "icon-192.png", "icon-512.png"]))); self.skipWaiting(); });
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", e => { if (new URL(e.request.url).origin !== location.origin) return; e.respondWith(fetch(e.request).catch(() => caches.match(e.request))); });
