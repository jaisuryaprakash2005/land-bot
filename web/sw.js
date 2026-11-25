// Service worker: lets the phone install the app and opens the screen even on a weak connection.
// Always tries the laptop first; falls back to the last saved copy of the screen.
const CACHE = "landbot-v1";
const SHELL = ["./", "index.html", "style.css", "app.js", "manifest.json",
               "vendor/msal-browser.min.js", "icons/icon-192.png", "icons/icon-64.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  // Only the screen files; never cache AI calls, Microsoft calls or WhatsApp shares.
  if (e.request.method !== "GET" || url.origin !== location.origin || !url.pathname.startsWith("/app/")) return;
  e.respondWith(
    fetch(e.request)
      .then((res) => {
        if (res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(e.request, copy)); }
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
