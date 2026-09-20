// Bump this on every deploy that changes any cached file. It's the only
// thing that makes the service worker notice new code exists — forgetting
// to bump it just means the network-first strategy below re-fetches fresh
// files anyway, so this mainly controls when the OLD cache gets swept away.
const CACHE = "moss-cache-v5";

// App-shell files: always try the network first so an installed PWA picks
// up new code as soon as it's online, falling back to cache only when
// offline. (A pure cache-first strategy here would freeze an installed
// app on whatever version was cached at install time, forever.)
const SHELL_ASSETS = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./db.js",
  "./auth.js",
  "./graph.js",
  "./config.js",
  "./msal-browser.min.js",
  "./manifest.json"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(SHELL_ASSETS)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return; // don't intercept cross-origin (e.g. future Graph API calls)

  event.respondWith(
    fetch(event.request)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((cache) => cache.put(event.request, copy)).catch(() => {});
        return res;
      })
      .catch(() =>
        caches.match(event.request).then((cached) => {
          if (cached) return cached;
          // Never cached and offline: fall back to the shell itself for a
          // navigation request so the app still opens; otherwise let it fail.
          if (event.request.mode === "navigate") {
            return caches.match("./index.html");
          }
          return Response.error();
        })
      )
  );
});
