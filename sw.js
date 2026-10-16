const CACHE = "falaj-5523d73b59";
const CORE = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png"];
// cache:"reload" = تُجلب ملفات الإصدار الجديد من الخادم مباشرة وليس من ذاكرة المتصفح المؤقتة
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE.map(u => new Request(u, { cache: "reload" }))))); self.skipWaiting(); });
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE && k !== "falaj-share").map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  // ملف مُشارك للتطبيق من واتساب: يُحفظ ثم يُفتح التطبيق ليستلمه
  if (e.request.method === "POST" && url.pathname.endsWith("/share-target")) {
    e.respondWith((async () => {
      try {
        const form = await e.request.formData(), file = form.get("file");
        if (file) {
          const c = await caches.open("falaj-share");
          await c.put(new URL("shared-file", self.registration.scope).href, new Response(file));
        }
      } catch (err) {}
      return Response.redirect(new URL("./?shared=1", self.registration.scope).href, 303);
    })());
    return;
  }
  if (e.request.method !== "GET") return;
  // من الذاكرة فوراً (يعمل بدون إنترنت) مع تحديث في الخلفية
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(hit => {
    const net = fetch(e.request).then(res => {
      if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
      return res;
    }).catch(() => hit);
    return hit || net;
  }));
});
