const CACHE = "ahliya-4e70b554bd";
const CORE = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE))); self.skipWaiting(); });
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE && k !== "ahliya-share").map(k => caches.delete(k)))));
  self.clients.claim();
});
// serve from cache instantly (works offline), refresh the cache in the background
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  // a file shared to the app (WhatsApp → مشاركة → الجمعيات): park it, then open the app which merges it
  if (e.request.method === "POST" && url.pathname.endsWith("/share-target")) {
    e.respondWith((async () => {
      try {
        const form = await e.request.formData(), file = form.get("file");
        if (file) {
          const c = await caches.open("ahliya-share");
          await c.put(new URL("shared-file", self.registration.scope).href,
            new Response(file, { headers: { "x-name": encodeURIComponent(file.name || "") } }));
        }
      } catch (err) {}
      return Response.redirect(new URL("./?shared=1#aza/sync", self.registration.scope).href, 303);
    })());
    return;
  }
  if (e.request.method !== "GET") return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(hit => {
    const net = fetch(e.request).then(res => {
      if (res.ok || res.type === "opaque") { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
      return res;
    }).catch(() => hit);
    return hit || net;
  }));
});
