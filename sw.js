// v4: red primero (siempre la última versión después de un deploy) y caché como respaldo sin conexión.
const CACHE_NAME = "stock-buscador-v4";
const SHELL = ["/", "/app.js", "/xlsx-worker.js", "/manifest.webmanifest", "/icon.svg"];
// Datos que se guardan para poder consultar el stock aunque se caiga la red.
const OFFLINE_API = ["/api/stock-data", "/api/current", "/api/history", "/api/version"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;
  const isApi = url.pathname.startsWith("/api/");
  if (isApi && !OFFLINE_API.includes(url.pathname)) return;

  const cacheKey = isApi || request.mode === "navigate" ? url.pathname : request;
  event.respondWith(
    fetch(request)
      .then((response) => {
        if (response.ok) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(cacheKey, copy));
        }
        return response;
      })
      .catch(() => caches.match(cacheKey).then((cached) => cached || caches.match("/")))
  );
});
