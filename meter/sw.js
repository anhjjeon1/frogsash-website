// 청개구리샤시 소음측정기 — 오프라인 열기용 서비스워커
//  · 앱 화면(index.html)은 "인터넷 먼저": 인터넷이 되면 늘 최신판, 안 되면(3.5초 안에 응답 없으면) 저장해 둔 판
//  · 아이콘·매니페스트는 저장본 먼저
//  · 날씨·지도·AI 등 다른 사이트 요청은 건드리지 않음(인터넷 없으면 그 기능만 안 됨)
const CACHE = "frog-meter-v1";
const CORE = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png", "./icon-512-mask.png", "./apple-touch-icon.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k.startsWith("frog-meter-") && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

function timeout(ms) { return new Promise((_, rej) => setTimeout(() => rej(new Error("timeout")), ms)); }

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (!url.pathname.startsWith(new URL("./", self.location).pathname)) return;

  if (req.mode === "navigate" || url.pathname.endsWith("/") || url.pathname.endsWith("/index.html")) {
    e.respondWith((async () => {
      const cache = await caches.open(CACHE);
      try {
        const res = await Promise.race([fetch(req, { cache: "no-store" }), timeout(3500)]);
        if (res && res.ok) { cache.put("./index.html", res.clone()); return res; }
        throw new Error("bad");
      } catch (err) {
        const hit = await cache.match("./index.html") || await cache.match("./");
        return hit || Response.error();
      }
    })());
    return;
  }

  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {
    if (res && res.ok) { const cp = res.clone(); caches.open(CACHE).then(c => c.put(req, cp)); }
    return res;
  })));
});
