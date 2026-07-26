/**
 * 네이버 뉴스 검색 API 중계 서버 (Cloudflare Workers)
 * ────────────────────────────────────────────────────────────
 * 휴대폰 브라우저는 네이버 API를 직접 호출할 수 없다.
 *   1) 응답에 CORS 헤더가 없어 브라우저가 차단하고,
 *   2) Client Secret 을 HTML 안에 넣으면 누구나 훔쳐 쓸 수 있다.
 * 이 워커가 그 사이에서 키를 숨긴 채 대신 호출해 준다.
 *
 * 배포 방법은 stock/README.md 참고.
 *
 * 필수 시크릿 : NAVER_CLIENT_ID, NAVER_CLIENT_SECRET
 * 선택 변수   : ALLOW_ORIGIN  (예: https://two-hee.github.io / 미설정 시 * )
 *
 * 호출 예 :
 *   GET https://<워커주소>/?ping                      → 상태 확인
 *   GET https://<워커주소>/?query=삼성전자&display=100&start=1&sort=date
 */

const NAVER_ENDPOINT = 'https://openapi.naver.com/v1/search/news.json';

export default {
  async fetch(request, env) {
    const cors = {
      'Access-Control-Allow-Origin': env.ALLOW_ORIGIN || '*',
      'Access-Control-Allow-Methods': 'GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '86400',
      'Vary': 'Origin'
    };

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors });
    if (request.method !== 'GET')     return json({ error: 'GET 요청만 지원합니다.' }, 405, cors);

    const url = new URL(request.url);

    // 연결 확인용 (키가 없어도 응답)
    if (url.searchParams.has('ping')) {
      return json({
        ok: true,
        service: 'naver-news-proxy',
        configured: Boolean(env.NAVER_CLIENT_ID && env.NAVER_CLIENT_SECRET)
      }, 200, cors);
    }

    const query = (url.searchParams.get('query') || '').trim();
    if (!query) return json({ error: 'query 파라미터가 필요합니다.' }, 400, cors);

    if (!env.NAVER_CLIENT_ID || !env.NAVER_CLIENT_SECRET) {
      return json({ error: '워커에 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 시크릿이 설정되지 않았습니다.' }, 500, cors);
    }

    const display = clamp(parseInt(url.searchParams.get('display'), 10) || 100, 1, 100);
    const start   = clamp(parseInt(url.searchParams.get('start'), 10) || 1, 1, 1000);
    const sort    = url.searchParams.get('sort') === 'sim' ? 'sim' : 'date';

    const target = `${NAVER_ENDPOINT}?query=${encodeURIComponent(query)}` +
                   `&display=${display}&start=${start}&sort=${sort}`;

    let res;
    try {
      res = await fetch(target, {
        headers: {
          'X-Naver-Client-Id': env.NAVER_CLIENT_ID,
          'X-Naver-Client-Secret': env.NAVER_CLIENT_SECRET
        }
      });
    } catch (err) {
      return json({ error: '네이버 API 호출 실패: ' + err.message }, 502, cors);
    }

    const body = await res.text();
    return new Response(body, {
      status: res.status,
      headers: {
        ...cors,
        'Content-Type': 'application/json; charset=utf-8',
        // 같은 질의를 반복해도 일일 호출 한도(25,000회)를 아끼도록 5분 캐시
        'Cache-Control': 'public, max-age=300'
      }
    });
  }
};

function clamp(n, min, max) { return Math.min(max, Math.max(min, n)); }

function json(obj, status, cors) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { ...cors, 'Content-Type': 'application/json; charset=utf-8' }
  });
}
