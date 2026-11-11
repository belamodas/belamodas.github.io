const CACHE_NAME = "bela-modas-vendedor-v1";
const APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "../sync.js",
  "../icone-avatar-app.svg",
  "../isabela-avatar.jpg"
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
        keys.filter(key => key.startsWith("bela-modas-vendedor-") && key !== CACHE_NAME)
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

  if (request.mode === "navigate") {
    const cachePromise = caches.open(CACHE_NAME);
    const network = fetch(request).then(response => {
      if (response && response.ok) cachePromise.then(cache => cache.put(request, response.clone()));
      return response;
    });
    event.respondWith((async () => {
      const cache = await cachePromise;
      return (await cache.match(request)) || (await network.catch(() => cache.match("./index.html"))) || Response.error();
    })());
    return;
  }

  const isAppFile = APP_SHELL.some(path => new URL(path, self.location.href).pathname === url.pathname);
  if (!isAppFile) return;
  event.respondWith(
    caches.match(request).then(cached => cached || fetch(request).then(response => {
      if (response && response.ok) caches.open(CACHE_NAME).then(cache => cache.put(request, response.clone()));
      return response;
    }))
  );
});
