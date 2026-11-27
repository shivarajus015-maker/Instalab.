const CACHE_NAME = "instalab-shell-v2";
const STATIC_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./app-config.js",
  "./assets/icon-192.png",
  "./assets/icon-512.png",
  "./assets/logo-kn.png",
  "./assets/logo-en.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(STATIC_ASSETS).catch(() => {});
    })
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);

  // Apps Script API calls must always hit the network so auth & latest sheet data stay fresh
  if (url.hostname.includes("script.google.com") || url.searchParams.has("action")) {
    return;
  }

  // Same-origin static assets: cache-first with network fallback
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(req).then((cachedRes) => {
        if (cachedRes) return cachedRes;
        return fetch(req).then((networkRes) => {
          if (networkRes && networkRes.status === 200) {
            const clone = networkRes.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, clone)).catch(() => {});
          }
          return networkRes;
        });
      })
    );
  }
});
