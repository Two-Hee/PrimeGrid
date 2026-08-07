/* PrimeGrid — 공통 UI 유틸 (데이터 로딩, 숫자 서식, SVG 차트) */
(function (global) {
  'use strict';

  var DATA_URL = (document.currentScript && document.currentScript.dataset.root ? document.currentScript.dataset.root : '') + 'data/poa.json';
  var cache = null;

  function load() {
    if (cache) return Promise.resolve(cache);
    return fetch(DATA_URL).then(function (r) {
      if (!r.ok) throw new Error('데이터를 불러오지 못했습니다 (' + r.status + ')');
      return r.json();
    }).then(function (d) { cache = d; return d; });
  }

  function num(v, dp) {
    if (v === null || v === undefined || !isFinite(v)) return '—';
    return v.toLocaleString('ko-KR', { minimumFractionDigits: dp || 0, maximumFractionDigits: dp || 0 });
  }
  function pct(v, dp) {
    if (v === null || v === undefined || !isFinite(v)) return '—';
    return (v * 100).toFixed(dp === undefined ? 1 : dp) + '%';
  }
  function signPct(v, dp) {
    if (v === null || v === undefined || !isFinite(v)) return '—';
    return (v >= 0 ? '+' : '') + (v * 100).toFixed(dp === undefined ? 1 : dp) + '%';
  }

  /** 지점 선택 <select> 채우기 (자료 12개월 지점 우선) */
  function fillStations(sel, stations, initial) {
    var full = stations.filter(function (s) { return s.vm === 12; });
    var part = stations.filter(function (s) { return s.vm < 12; });
    sel.innerHTML = '';
    function group(label, list) {
      if (!list.length) return;
      var g = document.createElement('optgroup');
      g.label = label;
      list.forEach(function (s) {
        var o = document.createElement('option');
        o.value = s.n;
        o.textContent = s.n + (s.vm < 12 ? ' (자료 ' + s.vm + '개월)' : '');
        g.appendChild(o);
      });
      sel.appendChild(g);
    }
    group('관측자료 12개월 완비 (' + full.length + '개 지점)', full);
    group('부분 자료 — 참고용 (' + part.length + '개 지점)', part);
    sel.value = initial && stations.some(function (s) { return s.n === initial; }) ? initial : '서울';
  }

  var NS = 'http://www.w3.org/2000/svg';
  function el(tag, attrs, text) {
    var e = document.createElementNS(NS, tag);
    for (var k in attrs) if (attrs[k] !== null && attrs[k] !== undefined) e.setAttribute(k, attrs[k]);
    if (text !== undefined) e.textContent = text;
    return e;
  }
  function css(name) {
    return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  }

  /**
   * 월별 그룹 막대 차트.
   * @param svg      대상 <svg>
   * @param labels   x축 라벨
   * @param series   [{name, values:[], color}]
   * @param opt      {unit, band:[{lo,hi}] 정상범위 밴드, dp}
   */
  function barChart(svg, labels, series, opt) {
    opt = opt || {};
    var W = 720, H = 260, P = { t: 16, r: 12, b: 34, l: 52 };
    svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H);
    svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
    while (svg.firstChild) svg.removeChild(svg.firstChild);

    var max = 0;
    series.forEach(function (s) { s.values.forEach(function (v) { if (isFinite(v) && v > max) max = v; }); });
    if (opt.band) opt.band.forEach(function (b) { if (b && b.hi > max) max = b.hi; });
    if (max <= 0) max = 1;
    var step = Math.pow(10, Math.floor(Math.log10(max))) / 2;
    max = Math.ceil(max / step) * step;

    var iw = W - P.l - P.r, ih = H - P.t - P.b;
    var y = function (v) { return P.t + ih - (v / max) * ih; };
    var bw = iw / labels.length;

    var gridC = css('--border'), dimC = css('--text-dim');

    // y축 눈금
    for (var i = 0; i <= 4; i++) {
      var v = max * i / 4;
      svg.appendChild(el('line', { x1: P.l, x2: W - P.r, y1: y(v), y2: y(v), stroke: gridC, 'stroke-width': 1 }));
      svg.appendChild(el('text', {
        x: P.l - 8, y: y(v) + 4, 'text-anchor': 'end',
        fill: dimC, 'font-size': 11, 'font-family': 'inherit'
      }, num(v, opt.dp === undefined ? 0 : opt.dp)));
    }

    // 정상범위 밴드
    if (opt.band) {
      opt.band.forEach(function (b, i) {
        if (!b) return;
        svg.appendChild(el('rect', {
          x: P.l + i * bw + bw * 0.1, y: y(b.hi), width: bw * 0.8, height: Math.max(1, y(b.lo) - y(b.hi)),
          fill: css('--ok'), opacity: 0.14, rx: 2
        }));
      });
    }

    // 막대
    var n = series.length, pad = bw * 0.18, gw = (bw - pad * 2) / n;
    series.forEach(function (s, si) {
      s.values.forEach(function (v, i) {
        if (!isFinite(v)) return;
        var h = Math.max(0, P.t + ih - y(v));
        svg.appendChild(el('rect', {
          x: P.l + i * bw + pad + si * gw, y: y(v), width: Math.max(1, gw - 1.5), height: h,
          fill: s.color, rx: 2
        }));
      });
    });

    // x축 라벨
    labels.forEach(function (l, i) {
      svg.appendChild(el('text', {
        x: P.l + i * bw + bw / 2, y: H - 12, 'text-anchor': 'middle',
        fill: dimC, 'font-size': 11, 'font-family': 'inherit'
      }, l));
    });
    svg.appendChild(el('line', { x1: P.l, x2: W - P.r, y1: P.t + ih, y2: P.t + ih, stroke: css('--border-strong'), 'stroke-width': 1 }));
  }

  global.UI = {
    load: load, num: num, pct: pct, signPct: signPct,
    fillStations: fillStations, barChart: barChart, css: css
  };
})(window);
