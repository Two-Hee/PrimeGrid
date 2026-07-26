/**
 * 네이버 뉴스 검색 API 중계 서버 (PC 로컬 테스트용)
 * ────────────────────────────────────────────────────────────
 * Cloudflare 배포 전에 내 PC 에서 먼저 확인해 보고 싶을 때 사용한다.
 * 외부 패키지 없이 Node 18 이상이면 바로 실행된다.
 *
 *   NAVER_CLIENT_ID=발급받은아이디 NAVER_CLIENT_SECRET=발급받은시크릿 \
 *   node stock/proxy/local-server.js
 *
 *   → http://localhost:8787 이 주소를 앱의 [API 설정] 칸에 넣으면 된다.
 *     (앱도 같은 PC 의 브라우저에서 열어야 한다)
 */

const http = require('http');

const PORT = process.env.PORT || 8787;
const ID = process.env.NAVER_CLIENT_ID || '';
const SECRET = process.env.NAVER_CLIENT_SECRET || '';
const NAVER_ENDPOINT = 'https://openapi.naver.com/v1/search/news.json';

const clamp = (n, min, max) => Math.min(max, Math.max(min, n));

const server = http.createServer(async (req, res) => {
  const cors = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
  const send = (status, obj) =>
    res.writeHead(status, { ...cors, 'Content-Type': 'application/json; charset=utf-8' })
       .end(typeof obj === 'string' ? obj : JSON.stringify(obj));

  if (req.method === 'OPTIONS') return res.writeHead(204, cors).end();
  if (req.method !== 'GET') return send(405, { error: 'GET 요청만 지원합니다.' });

  const url = new URL(req.url, `http://${req.headers.host}`);

  if (url.searchParams.has('ping'))
    return send(200, { ok: true, service: 'naver-news-proxy(local)', configured: Boolean(ID && SECRET) });

  const query = (url.searchParams.get('query') || '').trim();
  if (!query) return send(400, { error: 'query 파라미터가 필요합니다.' });
  if (!ID || !SECRET) return send(500, { error: 'NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 환경변수를 설정하세요.' });

  const display = clamp(parseInt(url.searchParams.get('display'), 10) || 100, 1, 100);
  const start = clamp(parseInt(url.searchParams.get('start'), 10) || 1, 1, 1000);
  const sort = url.searchParams.get('sort') === 'sim' ? 'sim' : 'date';

  try {
    const r = await fetch(
      `${NAVER_ENDPOINT}?query=${encodeURIComponent(query)}&display=${display}&start=${start}&sort=${sort}`,
      { headers: { 'X-Naver-Client-Id': ID, 'X-Naver-Client-Secret': SECRET } }
    );
    send(r.status, await r.text());
  } catch (err) {
    send(502, { error: '네이버 API 호출 실패: ' + err.message });
  }
});

server.listen(PORT, () => {
  console.log(`네이버 뉴스 중계 서버 실행 중 → http://localhost:${PORT}`);
  if (!ID || !SECRET) console.log('※ NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 환경변수가 비어 있습니다.');
});
