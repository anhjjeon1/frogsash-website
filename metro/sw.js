var CACHE_NAME = 'metro-v15a';
var URLS_TO_CACHE = [
  '/metro/',
  '/metro/index.html',
  '/metro/manifest.json',
  '/metro/icon-192.png',
  '/metro/icon-512.png',
  'https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700;800;900&family=Noto+Sans+KR:wght@300;400;500;700;900&family=JetBrains+Mono:wght@400;500&display=swap',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
  'https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js'
];

self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(URLS_TO_CACHE);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(names) {
      return Promise.all(
        names.filter(function(n) { return n !== CACHE_NAME; })
          .map(function(n) { return caches.delete(n); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(e) {
  // POST 등 비-GET 요청은 서비스워커가 개입하지 않음 (GAS 사진 업로드 등)
  if (e.request.method !== 'GET') return;
  var url = e.request.url;
  // API 요청은 서비스워커가 개입하지 않음 (브라우저가 직접 처리)
  if (url.indexOf('script.google.com') >= 0 || url.indexOf('drive.google.com') >= 0) {
    return;
  }
  // 나머지는 캐시 우선, 실패 시 네트워크
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      return cached || fetch(e.request).then(function(resp) {
        if (resp.status === 200) {
          var clone = resp.clone();
          caches.open(CACHE_NAME).then(function(cache) { cache.put(e.request, clone); });
        }
        return resp;
      });
    }).catch(function() {
      if (e.request.destination === 'document') {
        return caches.match('/metro/index.html');
      }
    })
  );
});
