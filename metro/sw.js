// Metro SW v17.0 — 캐시 전체 삭제 + 네트워크 전용 모드
// 이전 SW 캐시 문제 근본 해결: 캐시를 사용하지 않음

self.addEventListener('install', function(e) {
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    // 모든 캐시 삭제
    caches.keys().then(function(names) {
      return Promise.all(names.map(function(n) { return caches.delete(n); }));
    }).then(function() {
      return self.clients.claim();
    }).then(function() {
      // 열려있는 모든 탭 강제 새로고침
      return self.clients.matchAll({type: 'window'});
    }).then(function(clients) {
      clients.forEach(function(client) {
        client.navigate(client.url);
      });
    })
  );
});

self.addEventListener('fetch', function(e) {
  // 서비스워커가 어떤 요청에도 개입하지 않음 — 브라우저가 직접 처리
  return;
});
