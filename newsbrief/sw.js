/**
 * 종목 뉴스 브리핑 – 서비스 워커
 * ────────────────────────────────────────────────────────────
 * 앱 화면과 PDF 라이브러리를 저장해 두어 아이콘 실행 시 즉시 뜨고,
 * 네트워크가 끊긴 곳에서도 화면은 열린다. (뉴스 수집은 인터넷 필요)
 * 뉴스 응답은 캐시하지 않는다 — 항상 최신 기사를 받아야 한다.
 */

const VERSION = 'v1';
const CACHE = 'newsbrief-' + VERSION;

const SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './lib/jspdf.umd.min.js',
  './lib/html2canvas.min.js',
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
    await Promise.all(SHELL.map(url =>
      cache.add(new Request(url, { cache: 'reload' })).catch(() => {})
    ));
    self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k.startsWith('newsbrief-') && k !== CACHE)
                          .map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;   // 뉴스·중계 서버 요청은 그대로 통과

  const isDoc = req.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname.endsWith('/');

  if (isDoc) {
    // 화면은 네트워크 우선 — 앱을 고치면 바로 반영되고, 끊기면 캐시로 뜬다
    event.respondWith((async () => {
      try {
        const res = await fetch(req);
        (await caches.open(CACHE)).put(req, res.clone());
        return res;
      } catch (_) {
        return (await caches.match(req)) ||
               (await caches.match('./index.html')) ||
               new Response('오프라인 상태이고 저장된 화면이 없습니다.', {
                 status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' }
               });
      }
    })());
    return;
  }

  // 라이브러리·아이콘은 캐시 우선
  event.respondWith((async () => {
    const hit = await caches.match(req);
    if (hit) return hit;
    const res = await fetch(req);
    (await caches.open(CACHE)).put(req, res.clone());
    return res;
  })());
});
