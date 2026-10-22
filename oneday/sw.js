// 원데이 견적 앱: 늘 최신 화면을 먼저 받고, 인터넷이 끊기면 저장해 둔 화면을 보여 줌
var CACHE = 'oneday-v11';
self.addEventListener('install', function (e) { self.skipWaiting(); });
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;   // 서버(GAS) 요청은 건드리지 않음
  e.respondWith(fetch(e.request).then(function (r) {
    var c = r.clone(); caches.open(CACHE).then(function (ca) { ca.put(e.request, c); });
    return r;
  }).catch(function () { return caches.match(e.request); }));
});
