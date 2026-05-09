// Metro SW v23.26 - 자가 제거 + 강제 새로고침 (2026-05-09 16:30)
// 옛 SW가 v23.23 HTML을 캐시해서 반환하던 사고 해결: SW 자체를 한 번에 제거하고
// 클라이언트를 cache-buster 쿼리스트링과 함께 강제 재로드.

self.addEventListener('install', function(e) {
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil((async function() {
    try {
      // 1. 모든 캐시 삭제
      var names = await caches.keys();
      await Promise.all(names.map(function(n) { return caches.delete(n); }));

      // 2. 클라이언트 목록 확보 (unregister 전에)
      var clients = await self.clients.matchAll({ type: 'window' });

      // 3. 이 SW 자체를 등록 해제
      await self.registration.unregister();

      // 4. 모든 탭을 cache-buster 쿼리스트링과 함께 강제 재로드
      var stamp = Date.now();
      clients.forEach(function(client) {
        try {
          var u = new URL(client.url);
          u.searchParams.set('_sw', stamp);
          client.navigate(u.toString());
        } catch (e) {
          client.navigate(client.url);
        }
      });
    } catch (err) {
      // activate 실패 시에도 최소한 SW 등록 해제 시도
      try { await self.registration.unregister(); } catch (e) {}
    }
  })());
});

self.addEventListener('fetch', function(e) {
  // 어떤 요청도 가로채지 않음 — 브라우저가 직접 처리
  return;
});
