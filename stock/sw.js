/**
 * 종목 뉴스 브리핑 – 서비스 워커
 * ────────────────────────────────────────────────────────────
 * · 앱 화면(HTML·아이콘·매니페스트)을 캐시해 두어 홈 화면 아이콘으로
 *   실행했을 때 즉시 뜨고, 지하철·엘리베이터처럼 네트워크가 끊긴 곳에서도
 *   화면은 열리게 한다. (뉴스 수집 자체는 당연히 인터넷이 필요하다)
 * · 뉴스 API 응답은 절대 캐시하지 않는다 — 항상 최신 기사를 받아야 한다.
 */

const VERSION = 'v1';
const CACHE = 'stock-briefing-' + VERSION;

const SHELL = [
  './naver_news_briefing.html',
  './news_briefing.html',
  './manifest.webmanifest',
  './manifest-rss.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-192.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png'
];

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    // 한 개가 실패해도 나머지는 캐시되도록 개별 처리
    await Promise.all(SHELL.map(url =>
      cache.add(new Request(url, { cache: 'reload' })).catch(() => {})
    ));
    self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k.startsWith('stock-briefing-') && k !== CACHE)
                          .map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;          // 뉴스 API·기사 링크는 그대로 통과

  const isDoc = req.mode === 'navigate' || url.pathname.endsWith('.html');

  if (isDoc) {
    // 화면은 네트워크 우선 — 앱을 고치면 바로 반영되고, 끊기면 캐시로 뜬다
    event.respondWith((async () => {
      try {
        const res = await fetch(req);
        const cache = await caches.open(CACHE);
        cache.put(req, res.clone());
        return res;
      } catch (_) {
        return (await caches.match(req)) ||
               (await caches.match('./naver_news_briefing.html')) ||
               new Response('오프라인 상태이고 저장된 화면이 없습니다.', {
                 status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' }
               });
      }
    })());
    return;
  }

  // 아이콘·매니페스트 등 정적 자원은 캐시 우선
  event.respondWith((async () => {
    const hit = await caches.match(req);
    if (hit) return hit;
    try {
      const res = await fetch(req);
      const cache = await caches.open(CACHE);
      cache.put(req, res.clone());
      return res;
    } catch (err) {
      throw err;
    }
  })());
});
