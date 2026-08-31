const CACHE = "beezness-v2";
const FILES = ["./", "./index.html", "./app.js", "./icon.png", "./manifest.json"];
self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
// index.html/app.js change on every deploy, so they go network-first (falling
// back to cache only when offline) — otherwise a stale cached copy would
// keep serving indefinitely even after a fresh version is pushed.
const NETWORK_FIRST_SUFFIXES = ["/", "/index.html", "/honey-till.html", "/app.js"];
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  const path = new URL(e.request.url).pathname;
  const networkFirst = NETWORK_FIRST_SUFFIXES.some((suffix) => path.endsWith(suffix));
  if (networkFirst) {
    e.respondWith(
      fetch(e.request).then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
        return res;
      }).catch(() => caches.match(e.request).then((hit) => hit || caches.match("./index.html")))
    );
    return;
  }
  e.respondWith(
    caches.match(e.request).then((hit) => hit || fetch(e.request).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
      return res;
    }))
  );
});