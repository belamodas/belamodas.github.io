const CACHE_NAME = "bela-modas-app-v10";
const APP_SHELL = [
  "./",
  "./index.html",
  "./sync.js",
  "./manifest.webmanifest",
  "./icone-bela.svg",
  "./isabela-avatar.jpg",
  "./icone-avatar-app.svg"
];

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(key => key.startsWith("bela-modas-app-") && key !== CACHE_NAME)
          .map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  const request = event.request;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;
  const isAppFile = APP_SHELL.some(path => new URL(path, self.location.href).pathname === url.pathname);
  if (request.mode !== "navigate" && !isAppFile) return;
  // Prioriza sempre o código publicado; usa o cache somente sem conexão.
  event.respondWith(
    fetch(request).then(response => {
      if (response && response.ok) {
        const copy = response.clone();
        event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.put(request, copy)));
      }
      return response;
    }).catch(async () => {
      const cache = await caches.open(CACHE_NAME);
      return (await cache.match(request)) ||
        (request.mode === "navigate" ? await cache.match("./index.html") : undefined) ||
        Response.error();
    })
  );
});
