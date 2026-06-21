// バージョンを上げるたびに古いキャッシュが自動削除される
const CACHE = 'jhm-v11';
const ASSETS = [
  '/japan-history/mobile/',
  '/japan-history/mobile/index.html',
  '/japan-history/mobile/nenpyo_icon.png',
  '/japan-history/japan_history_data.json',
  '/japan-history/japan.topojson'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  // 古いキャッシュを全て削除
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

// ネットワークファースト：常に最新を取得し、失敗時のみキャッシュ使用
self.addEventListener('fetch', e => {
  e.respondWith(
    fetch(e.request).then(res => {
      if (res.ok) {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
      }
      return res;
    }).catch(() => caches.match(e.request))
  );
});
