const CACHE_NAME = "bela-modas-app-v7";
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

  // Abre imediatamente a última página guardada e atualiza em segundo plano.
  if (request.mode === "navigate") {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_NAME);
      const paginaSalva = await cache.match("./index.html");
      const atualizacao = fetch(request).then(response => {
        if (response && response.ok) {
          cache.put("./index.html", response.clone());
        }
        return response;
      });
      if (paginaSalva) {
        event.waitUntil(atualizacao.then(() => undefined).catch(() => undefined));
        return paginaSalva;
      }
      return atualizacao.catch(() => cache.match("./"));
    })());
    return;
  }

  // Somente arquivos essenciais do aplicativo entram no cache.
  const isAppFile = APP_SHELL.some(path => new URL(path, self.location.href).pathname === url.pathname);
  if (!isAppFile) return;

  event.respondWith(
    caches.match(request).then(cached => {
      if (cached) return cached;
      return fetch(request).then(response => {
        if (response && response.ok) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(request, copy));
        }
        return response;
      });
    })
  );
});
