#!/usr/bin/env python3
"""여러 페이지를 CSS·JS·데이터까지 인라인한 단일 HTML 파일로 합친다.

각 페이지는 독립 실행을 전제로 만들어져 있어 id가 겹친다(st, beta, tbl, src …).
그래서 페이지마다 접두사를 붙여 충돌을 없앤다.

  - HTML의 id= / for= / aria-describedby= 속성
  - 페이지 스크립트의 $ 헬퍼 (접두사를 자동으로 붙이도록 재정의)
  - 스크립트 안의 document.getElementById('x') 와 '#x' 선택자

실행: python3 build/build_standalone.py
결과: dist/PrimeGrid_태양광분석.html
"""
import re, pathlib, html

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / 'dist' / 'PrimeGrid_태양광분석.html'

# (파일, 접두사, 탭 이름, 탭 설명)
PAGES = [
    ('stations.html',        's', '전국 지점 자료', '65개 관측지점 표와 지점별 월별 상세'),
    ('tools/yield.html',     'y', '예상 발전량',    '지역·설비 조건별 월별 POA와 30년 발전량'),
    ('tools/diagnosis.html', 'd', '발전량 진단',    '실제 발전량을 지역 기준선과 비교'),
    ('method.html',          'm', '데이터·계산방법', '원자료·수식·검증·한계 전문'),
]

def read(p):
    return (ROOT / p).read_text(encoding='utf-8')

def grab_main(src):
    """<main …> … </main> 안쪽을 꺼낸다."""
    m = re.search(r'<main[^>]*>(.*?)</main>', src, re.S)
    if not m:
        raise SystemExit('main 태그를 찾지 못했습니다')
    return m.group(1)

def grab_page_script(src):
    """페이지 고유 로직이 담긴 마지막 인라인 <script>를 꺼낸다."""
    blocks = re.findall(r'<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>', src, re.S)
    # <head>의 테마 선적용 스니펫은 제외
    blocks = [b for b in blocks if 'localStorage.getItem("pg-theme")' not in b]
    return blocks[-1] if blocks else ''

def grab_style(src):
    """페이지가 따로 가진 <style> 블록(있으면)을 꺼낸다."""
    m = re.search(r'<style[^>]*>(.*?)</style>', src, re.S)
    return m.group(1) if m else ''

# id를 건드리면 안 되는 것들 (문서 전역에서 하나만 존재)
GLOBAL_IDS = {'themeToggle'}

def prefix_html(frag, pfx):
    def sub(m):
        attr, quote, val = m.group(1), m.group(2), m.group(3)
        if val in GLOBAL_IDS:
            return m.group(0)
        return f'{attr}={quote}{pfx}-{val}{quote}'
    return re.sub(r'\b(id|for|aria-describedby|aria-labelledby)=(["\'])([A-Za-z][\w-]*)\2', sub, frag)

def prefix_js(js, pfx):
    # $ 헬퍼가 접두사를 자동으로 붙이게 한다 → $('x') 호출은 손대지 않아도 된다
    js = js.replace(
        "var $ = function (id) { return document.getElementById(id); };",
        f"var PFX = '{pfx}-';\n  var $ = function (id) {{ return document.getElementById(PFX + id); }};")
    # 동적으로 만들어 넣는 id/for 문자열 (진단 페이지의 월별 입력칸)
    js = js.replace("""id="m' + i + '\"""", f"""id="' + PFX + 'm' + i + '\"""")
    js = js.replace("""for="m' + i + '\"""", f"""for="' + PFX + 'm' + i + '\"""")
    # $ 를 쓰지 않고 직접 부르는 곳
    js = re.sub(r"document\.getElementById\((['\"])([A-Za-z][\w-]*)\1\)",
                lambda m: f"document.getElementById('{pfx}-{m.group(2)}')"
                          if m.group(2) not in GLOBAL_IDS else m.group(0), js)
    # querySelector/All 안의 #id 선택자
    def qs(m):
        head, quote, sel = m.group(1), m.group(2), m.group(3)
        new = re.sub(r'#([A-Za-z][\w-]*)',
                     lambda x: '#' + x.group(1) if x.group(1) in GLOBAL_IDS else f'#{pfx}-{x.group(1)}',
                     sel)
        return f'{head}({quote}{new}{quote}'
    js = re.sub(r'(querySelector(?:All)?)\((["\'])([^"\']*#[^"\']*)\2', qs, js)
    return js

def strip_nav_and_print(frag):
    """단일 파일에서는 페이지 이동 링크 대신 탭을 쓰므로 내부 링크를 탭 전환으로 바꾼다."""
    frag = re.sub(r'href="(?:\.\./)?tools/yield\.html"', 'href="#tab-y" data-tab="y"', frag)
    frag = re.sub(r'href="(?:\.\./)?tools/diagnosis\.html"', 'href="#tab-d" data-tab="d"', frag)
    frag = re.sub(r'href="(?:\.\./)?stations\.html"', 'href="#tab-s" data-tab="s"', frag)
    frag = re.sub(r'href="(?:\.\./)?method\.html(#[\w-]*)?"', 'href="#tab-m" data-tab="m"', frag)
    frag = re.sub(r'href="(?:\.\./)?index\.html"', 'href="#tab-s" data-tab="s"', frag)
    # 파일로 배포하므로 내부 파일 링크는 의미가 없다 — 링크를 벗겨 텍스트만 남긴다
    frag = re.sub(r'<a href="(?:\.\./)?(?:data|assets)/[^"]+">([^<]*)</a>',
                  r'<span class="mono">\1</span>', frag)
    return frag

# ── 자산 읽기 ──
css = read('assets/style.css')
libs = '\n'.join(read(f) for f in
                 ('data/poa-data.js', 'data/verification-data.js',
                  'assets/poa.js', 'assets/ui.js'))

sections, scripts, extra_css = [], [], []
for path, pfx, label, desc in PAGES:
    src = read(path)
    extra_css.append(grab_style(src))
    frag = strip_nav_and_print(prefix_html(grab_main(src), pfx))
    sections.append(f'<section class="tabpane" id="tab-{pfx}" hidden>\n{frag}\n</section>')
    js = grab_page_script(src)
    if js.strip():
        scripts.append(f'/* ── {label} ── */\n' + prefix_js(js, pfx))

tabs = '\n'.join(
    f'      <button class="tabbtn" data-tab="{p}" type="button">{l}</button>'
    for _, p, l, _ in PAGES)

FAVICON = ("data:image/svg+xml,%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 viewBox=%270 0 32 32%27%3E"
           "%3Ccircle cx=%2716%27 cy=%2716%27 r=%277%27 fill=%27%23e08a1e%27/%3E"
           "%3Cg stroke=%27%23e08a1e%27 stroke-width=%272.6%27 stroke-linecap=%27round%27%3E"
           "%3Cpath d=%27M16 2v4M16 26v4M2 16h4M26 16h4M6 6l3 3M23 23l3 3M26 6l-3 3M9 23l-3 3%27/%3E"
           "%3C/g%3E%3C/svg%3E")

doc = f'''<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PrimeGrid 태양광 발전량 분석 — 단일 파일판</title>
<meta name="description" content="기상청 ASOS 10년 일사량 관측자료 기반 태양광 분석. 인터넷 없이 이 파일 하나로 동작합니다.">
<link rel="icon" href="{FAVICON}">
<script>(function(){{try{{var t=localStorage.getItem("pg-theme");if(t)document.documentElement.setAttribute("data-theme",t);}}catch(e){{}}}})();</script>
<style>
{css}
{"".join(extra_css)}
/* ── 단일 파일판 전용 ── */
.tabbar {{ display: flex; gap: 4px; flex-wrap: wrap; margin-left: auto; }}
.tabbtn {{
  font: inherit; font-size: .92rem; font-weight: 500; cursor: pointer;
  padding: 6px 12px; border-radius: 7px; border: 1px solid transparent;
  background: transparent; color: var(--text-mid);
}}
.tabbtn:hover {{ color: var(--text); background: var(--bg-soft); border-color: var(--border); }}
.tabbtn[aria-selected="true"] {{
  color: var(--text); background: var(--bg-soft);
  border-color: var(--border); box-shadow: inset 0 -2px 0 var(--accent);
}}
.tabpane[hidden] {{ display: none; }}
.offline-note {{
  background: var(--ok-soft); border: 1px solid transparent; border-left: 3px solid var(--ok);
  border-radius: 0 7px 7px 0; padding: 10px 16px; margin: 16px auto 0; max-width: 1080px;
  font-size: .85rem; color: var(--text-mid);
}}
@media print {{
  .tabpane[hidden] {{ display: none !important; }}
  .offline-note {{ display: none; }}
}}
</style>
</head>
<body>

<header class="site-head no-print">
  <div class="wrap">
    <a class="logo" href="#tab-s" data-tab="s"><span class="dot"></span>PrimeGrid</a>
    <nav class="tabbar" role="tablist">
{tabs}
      <button class="theme-btn" id="themeToggle" type="button">☾</button>
    </nav>
  </div>
</header>

<div class="offline-note no-print">
  <strong>단일 파일판.</strong> 자료·계산식·서식이 이 파일 하나에 모두 들어 있습니다.
  인터넷이 없어도, 다른 파일이 없어도 그대로 동작합니다.
  기상청 ASOS 일사량 관측자료 2016–2025년 177,818건 · 65개 지점 기준.
</div>

<main class="wrap" style="padding:0">
{chr(10).join(sections)}
</main>

<footer class="site-foot">
  <div class="wrap">
    <div>자료 출처: 기상청 기상자료개방포털 ASOS 일사량 관측자료 (2016–2025)</div>
    <div>PrimeGrid · 단일 파일판</div>
  </div>
  <div class="foot-notice">
    <p><strong>저작권.</strong> 본 자료에 담긴 분석 내용, 계산 방법, 산출 자료 및 화면 구성의
    저작권은 PrimeGrid 운영자에게 있습니다. 무단 복제·배포·상업적 이용을 금합니다.
    원자료인 기상청 ASOS 관측자료의 저작권은 기상청에 있습니다.</p>
    <p style="margin-bottom:0"><strong>정확도.</strong> 게시된 모든 수치는 관측 기후값과 계산 모델에 따른
    <strong>추정값이며 오차가 있습니다.</strong> 음영·오염·적설·모듈온도·출력제한 등은 반영되어 있지 않아
    실제 발전량과 차이가 날 수 있습니다. 본 자료는 참고용이며 발전소 설계·감정·금융 판단을
    대체하지 않습니다. 이용자의 의사결정에 따른 결과에 대해 운영자는 책임을 지지 않습니다.</p>
  </div>
</footer>

<script>
{libs}
</script>

<script>
/* 탭 전환 */
(function () {{
  'use strict';
  var panes = [].slice.call(document.querySelectorAll('.tabpane'));
  var btns  = [].slice.call(document.querySelectorAll('.tabbtn'));
  function show(t) {{
    panes.forEach(function (p) {{ p.hidden = (p.id !== 'tab-' + t); }});
    btns.forEach(function (b) {{ b.setAttribute('aria-selected', String(b.dataset.tab === t)); }});
    try {{ history.replaceState(null, '', '#tab-' + t); }} catch (e) {{}}
    window.scrollTo({{ top: 0, behavior: 'instant' }});
  }}
  document.addEventListener('click', function (e) {{
    var el = e.target.closest('[data-tab]');
    if (!el) return;
    e.preventDefault();
    show(el.dataset.tab);
  }});
  var init = (location.hash.match(/^#tab-(\\w)$/) || [])[1] || '{PAGES[0][1]}';
  show(panes.some(function (p) {{ return p.id === 'tab-' + init; }}) ? init : '{PAGES[0][1]}');
  window.__pgShowTab = show;
}})();
</script>

<script>
{chr(10).join(scripts)}
</script>

<script>
/* 테마 토글 */
(function () {{
  'use strict';
  var KEY = 'pg-theme', btn = document.getElementById('themeToggle');
  if (!btn) return;
  function cur() {{
    return document.documentElement.getAttribute('data-theme') ||
      (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  }}
  function paint() {{
    var dark = cur() === 'dark';
    btn.textContent = dark ? '☀' : '☾';
    btn.title = dark ? '밝은 화면으로 전환' : '어두운 화면으로 전환';
    btn.setAttribute('aria-label', btn.title);
  }}
  btn.addEventListener('click', function () {{
    var next = cur() === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    try {{ localStorage.setItem(KEY, next); }} catch (e) {{}}
    paint();
  }});
  paint();
}})();
</script>
</body>
</html>
'''

OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text(doc, encoding='utf-8')
print(f'{OUT}  ({OUT.stat().st_size:,} bytes)')
