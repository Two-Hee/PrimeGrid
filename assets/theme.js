/* PrimeGrid — 화면 모드 토글
 *
 * 기본값은 OS 설정(prefers-color-scheme)을 그대로 따릅니다.
 * 사용자가 토글을 누르면 그 선택을 localStorage에 저장하고
 * <html data-theme="light|dark">로 고정합니다.
 *
 * 첫 페인트 전에 data-theme을 붙여야 화면이 번쩍이지 않으므로,
 * 저장값 적용은 각 페이지 <head>의 인라인 스크립트가 담당합니다.
 * 이 파일은 버튼 표시와 클릭 처리만 맡습니다.
 */
(function () {
  'use strict';
  var KEY = 'pg-theme';
  var btn = document.getElementById('themeToggle');
  if (!btn) return;

  function current() {
    var set = document.documentElement.getAttribute('data-theme');
    if (set) return set;
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }

  function paint() {
    var dark = current() === 'dark';
    btn.textContent = dark ? '☀' : '☾';
    btn.title = dark ? '밝은 화면으로 전환' : '어두운 화면으로 전환';
    btn.setAttribute('aria-label', btn.title);
  }

  btn.addEventListener('click', function () {
    var next = current() === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    try { localStorage.setItem(KEY, next); } catch (e) { /* 저장 불가 시 이번 세션만 적용 */ }
    paint();
  });

  // 사용자가 고정 선택을 하지 않았다면 OS 설정 변화를 따라간다
  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', function () {
    if (!document.documentElement.getAttribute('data-theme')) paint();
  });

  paint();
})();
