/**
 * GET /api/imweb/banner-widget.js
 * 아임웹 사이트에 주입되는 자기완결형 JavaScript
 *
 * Phase 2 (2026-04-22): 이미지 다중 업로드 → 자동 슬라이더
 *   - 이미지 1장: 정적 모달 (기존과 동일)
 *   - 이미지 2+장: 자동 슬라이더 (5초 전환, 점/화살표 네비, 스와이프, 호버 일시정지, 루프)
 *
 * 아임웹 측 설치: <script src=".../api/imweb/banner-widget.js" defer></script>
 *
 * Content-Type: application/javascript
 * Cache-Control: 1시간
 */

import { NextResponse } from 'next/server';

const BASE_URL = process.env.NEXT_PUBLIC_APP_URL || 'https://app-eta-sandy-75.vercel.app';

const WIDGET_JS = `(function(){
  'use strict';
  if (window.__MAMORU_BANNER_LOADED) return;
  window.__MAMORU_BANNER_LOADED = true;

  var COOKIE_KEY = 'mamoru_banner_dismissed_at';
  var CONFIG_URL = '${BASE_URL}/api/imweb/banner-config';
  var Z_INDEX = 999999;
  var SLIDE_INTERVAL = 5000; // 5초 (잡스 권고)

  function getCookie(name) {
    var match = document.cookie.match(new RegExp('(^|;)\\\\s*' + name + '=([^;]*)'));
    return match ? decodeURIComponent(match[2]) : null;
  }

  function setCookie(name, value, hours) {
    var d = new Date();
    d.setTime(d.getTime() + (hours * 60 * 60 * 1000));
    document.cookie = name + '=' + encodeURIComponent(value) + ';expires=' + d.toUTCString() + ';path=/';
  }

  function shouldHide(config) {
    if (!config || !config.enabled) return true;
    var dismissed = getCookie(COOKIE_KEY);
    if (dismissed && config.dismiss_cookie_hours) {
      var elapsed = (Date.now() - parseInt(dismissed, 10)) / (1000 * 60 * 60);
      if (!isNaN(elapsed) && elapsed < config.dismiss_cookie_hours) return true;
    }
    return false;
  }

  function esc(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function injectStyles() {
    if (document.getElementById('mamoru-banner-styles')) return;
    var style = document.createElement('style');
    style.id = 'mamoru-banner-styles';
    style.textContent = [
      '#mamoru-banner-overlay{position:fixed;inset:0;z-index:' + Z_INDEX + ';background:rgba(26,26,26,0.72);display:flex;align-items:center;justify-content:center;padding:24px;animation:mamoru-fade-in 0.2s ease-out;}',
      '@keyframes mamoru-fade-in{from{opacity:0}to{opacity:1}}',
      '#mamoru-banner-modal{background:#FAF9F7;max-width:420px;width:100%;border-radius:8px;overflow:hidden;font-family:"Noto Sans KR",-apple-system,BlinkMacSystemFont,sans-serif;box-shadow:0 20px 60px rgba(0,0,0,0.3);position:relative;}',
      // 슬라이드 트랙
      '.mamoru-slide-track{position:relative;width:100%;overflow:hidden;background:#EEE;user-select:none;-webkit-user-select:none;touch-action:pan-y;}',
      '.mamoru-slide-list{display:flex;transition:transform 0.4s cubic-bezier(0.4,0,0.2,1);}',
      '.mamoru-slide-item{flex:0 0 100%;width:100%;display:block;}',
      '.mamoru-slide-item img{width:100%;display:block;pointer-events:none;}',
      '.mamoru-slide-item.clickable{cursor:pointer;}',
      '.mamoru-slide-item.clickable img{pointer-events:auto;}',
      // 화살표 (데스크톱만)
      '.mamoru-slide-arrow{position:absolute;top:50%;transform:translateY(-50%);width:34px;height:34px;border:none;background:rgba(26,26,26,0.5);color:#FFF;border-radius:50%;font-size:18px;cursor:pointer;display:none;align-items:center;justify-content:center;line-height:1;z-index:2;transition:background 0.15s;}',
      '.mamoru-slide-arrow:hover{background:rgba(26,26,26,0.75);}',
      '.mamoru-slide-arrow.prev{left:8px;}',
      '.mamoru-slide-arrow.next{right:8px;}',
      '@media (hover:hover) and (pointer:fine){.mamoru-slide-arrow{display:flex;}}',
      // 점 네비게이션
      '.mamoru-slide-dots{display:flex;gap:6px;justify-content:center;padding:8px 0;background:#FAF9F7;}',
      '.mamoru-slide-dot{width:7px;height:7px;border-radius:50%;border:none;background:#D0D0D0;cursor:pointer;padding:0;transition:all 0.2s;}',
      '.mamoru-slide-dot.active{background:#1A1A1A;width:20px;border-radius:4px;}',
      // 본문/액션
      '#mamoru-banner-body{padding:24px;}',
      '#mamoru-banner-title{font-size:17px;font-weight:700;color:#1A1A1A;margin:0 0 8px;letter-spacing:-0.3px;}',
      '#mamoru-banner-desc{font-size:14px;color:#555;margin:0;line-height:1.6;white-space:pre-line;}',
      '#mamoru-banner-actions{display:flex;gap:0;border-top:1px solid #EEE;}',
      '#mamoru-banner-actions button{flex:1;padding:14px 0;border:none;background:transparent;font-size:13px;color:#555;cursor:pointer;font-family:inherit;letter-spacing:-0.2px;}',
      '#mamoru-banner-actions button:hover{background:#F5F5F5;}',
      '#mamoru-banner-actions .primary{color:#1A1A1A;font-weight:700;border-left:1px solid #EEE;}'
    ].join('\\n');
    document.head.appendChild(style);
  }

  /** 슬라이드 마크업 생성 (1개면 정적, 2+개면 슬라이더) */
  function buildSlides(images) {
    var items = images.map(function(img) {
      var cls = img.link_url ? 'mamoru-slide-item clickable' : 'mamoru-slide-item';
      return '<div class="' + cls + '" data-link="' + esc(img.link_url || '') + '"><img src="' + esc(img.url) + '" alt=""></div>';
    }).join('');

    var html = '<div class="mamoru-slide-track"><div class="mamoru-slide-list">' + items + '</div>';
    if (images.length >= 2) {
      // 화살표
      html += '<button class="mamoru-slide-arrow prev" aria-label="이전">‹</button>';
      html += '<button class="mamoru-slide-arrow next" aria-label="다음">›</button>';
    }
    html += '</div>';

    // 점 네비게이션 (2+장만)
    if (images.length >= 2) {
      var dots = '';
      for (var i = 0; i < images.length; i++) {
        dots += '<button class="mamoru-slide-dot' + (i === 0 ? ' active' : '') + '" data-idx="' + i + '" aria-label="' + (i+1) + '번 슬라이드"></button>';
      }
      html += '<div class="mamoru-slide-dots">' + dots + '</div>';
    }

    return html;
  }

  /** 슬라이더 인터랙션 활성화 (이미지 2+장일 때) */
  function activateSlider(root, images) {
    if (images.length < 2) return;

    var list = root.querySelector('.mamoru-slide-list');
    var dots = root.querySelectorAll('.mamoru-slide-dot');
    var prevBtn = root.querySelector('.mamoru-slide-arrow.prev');
    var nextBtn = root.querySelector('.mamoru-slide-arrow.next');
    var track = root.querySelector('.mamoru-slide-track');
    var current = 0;
    var total = images.length;
    var timer = null;
    var paused = false;

    function render() {
      list.style.transform = 'translateX(-' + (current * 100) + '%)';
      dots.forEach(function(d, i) {
        d.classList.toggle('active', i === current);
      });
    }

    function goTo(idx) {
      current = ((idx % total) + total) % total;
      render();
    }

    function next() { goTo(current + 1); }
    function prev() { goTo(current - 1); }

    function startTimer() {
      stopTimer();
      if (paused) return;
      timer = setInterval(next, SLIDE_INTERVAL);
    }
    function stopTimer() {
      if (timer) { clearInterval(timer); timer = null; }
    }

    // 점 클릭
    dots.forEach(function(d) {
      d.addEventListener('click', function(e) {
        e.stopPropagation();
        var idx = parseInt(d.getAttribute('data-idx'), 10);
        goTo(idx);
        startTimer(); // 수동 이동 시 타이머 리셋
      });
    });

    // 화살표
    if (prevBtn) prevBtn.addEventListener('click', function(e) { e.stopPropagation(); prev(); startTimer(); });
    if (nextBtn) nextBtn.addEventListener('click', function(e) { e.stopPropagation(); next(); startTimer(); });

    // 호버 시 일시정지 (데스크톱)
    track.addEventListener('mouseenter', function() { paused = true; stopTimer(); });
    track.addEventListener('mouseleave', function() { paused = false; startTimer(); });

    // 모바일 스와이프
    var touchStartX = 0;
    var touchEndX = 0;
    track.addEventListener('touchstart', function(e) {
      touchStartX = e.changedTouches[0].screenX;
      paused = true; stopTimer();
    }, { passive: true });
    track.addEventListener('touchend', function(e) {
      touchEndX = e.changedTouches[0].screenX;
      var dx = touchEndX - touchStartX;
      if (Math.abs(dx) > 40) {
        if (dx < 0) next(); else prev();
      }
      paused = false;
      startTimer();
    }, { passive: true });

    startTimer();

    // 정리 훅 (overlay가 사라지면 타이머 해제)
    root.__cleanupSlider = stopTimer;
  }

  function render(config) {
    injectStyles();

    var images = (Array.isArray(config.images) && config.images.length > 0)
      ? config.images
      : (config.image_url ? [{ url: config.image_url, link_url: config.link_url || '' }] : []);

    if (images.length === 0) return; // 이미지 0장 = 배너 미표시

    var overlay = document.createElement('div');
    overlay.id = 'mamoru-banner-overlay';

    var slidesHtml = buildSlides(images);

    var hasText = config.title || config.description;
    var bodyHtml = hasText
      ? '<div id="mamoru-banner-body">' +
          (config.title ? '<h3 id="mamoru-banner-title">' + esc(config.title) + '</h3>' : '') +
          (config.description ? '<p id="mamoru-banner-desc">' + esc(config.description) + '</p>' : '') +
        '</div>'
      : '';

    overlay.innerHTML =
      '<div id="mamoru-banner-modal" role="dialog" aria-modal="true">' +
        slidesHtml +
        bodyHtml +
        '<div id="mamoru-banner-actions">' +
          '<button id="mamoru-banner-dismiss">오늘 하루 보지 않기</button>' +
          '<button id="mamoru-banner-close" class="primary">닫기</button>' +
        '</div>' +
      '</div>';

    function close() {
      if (overlay.__cleanupSlider) overlay.__cleanupSlider();
      if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
    }

    // "닫기" 버튼
    overlay.querySelector('#mamoru-banner-close').addEventListener('click', close);

    // "오늘 하루 보지 않기"
    overlay.querySelector('#mamoru-banner-dismiss').addEventListener('click', function() {
      setCookie(COOKIE_KEY, Date.now(), config.dismiss_cookie_hours || 24);
      close();
    });

    // 슬라이드 이미지 클릭 → 해당 슬라이드 link_url 열기
    overlay.querySelectorAll('.mamoru-slide-item.clickable').forEach(function(el) {
      el.addEventListener('click', function() {
        var href = el.getAttribute('data-link');
        if (href) window.open(href, '_blank', 'noopener');
      });
    });

    // 배경 클릭 시 닫기
    overlay.addEventListener('click', function(e) {
      if (e.target === overlay) close();
    });

    document.body.appendChild(overlay);

    // 슬라이더 활성화 (이미지 2+장일 때)
    activateSlider(overlay, images);
  }

  function init() {
    try {
      fetch(CONFIG_URL, { cache: 'no-store' })
        .then(function(r) { return r.json(); })
        .then(function(config) {
          if (shouldHide(config)) return;
          render(config);
        })
        .catch(function(e) {
          if (window.console) console.warn('[MAMORU banner] config 로드 실패', e);
        });
    } catch (e) {
      if (window.console) console.warn('[MAMORU banner] init 실패', e);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();

/* ===== MAMORU iframe 자동 높이조절 (전역) =====
   아임웹은 코드위젯 안의 <script>를 제거/미실행(특히 느린 회선)하므로,
   iframe 높이조절을 코드위젯에 의존하면 간헐적으로 콘텐츠가 잘린다.
   이 전역 스크립트(헤더 설치, 항상 실행)가 page.mamoru.kr 자식 iframe들을
   source 매칭으로 직접 사이징 → 코드위젯 <script> 실행 여부와 무관하게 안정.
   자식은 REQUEST_HEIGHT에 무조건 회신(force)하므로 초기 푸시를 놓쳐도 복구됨. (2026-06-29) */
(function(){
  if (window.__MAMORU_RESIZER_LOADED) return;
  window.__MAMORU_RESIZER_LOADED = true;
  var ORIGIN = 'https://page.mamoru.kr';
  function frames(){
    return [].slice.call(document.querySelectorAll('iframe')).filter(function(f){
      return (f.src || '').indexOf(ORIGIN) === 0;
    });
  }
  /* 아임웹 상단 고정 헤더(로고바) 아래 지점 — 모달을 그 아래에 배치해 가림 방지 */
  function topFixedBottom(){
    try {
      if (!document.elementsFromPoint) return 0;
      var pts = document.elementsFromPoint(Math.round(window.innerWidth / 2), 3);
      var b = 0, lim = window.innerHeight * 0.4;
      for (var i = 0; i < pts.length; i++){
        var el = pts[i];
        if (!el || el.tagName === 'IFRAME') continue;
        var cs = window.getComputedStyle(el);
        if (cs && (cs.position === 'fixed' || cs.position === 'sticky')){
          var rb = el.getBoundingClientRect().bottom;
          if (rb > b && rb <= lim) b = rb;   // 화면 상단 40% 내 고정요소만 헤더로 간주
        }
      }
      return b;
    } catch (_) { return 0; }
  }
  /* 모달 열림 동안 부모(아임웹) 페이지 스크롤 잠금 — 배경이 안 밀림 */
  var __mmLockY = null;
  function lockScroll(){
    if (__mmLockY !== null) return;
    __mmLockY = window.pageYOffset || document.documentElement.scrollTop || 0;
    var b = document.body;
    b.style.position = 'fixed'; b.style.top = (-__mmLockY) + 'px';
    b.style.left = '0'; b.style.right = '0'; b.style.width = '100%';
  }
  function unlockScroll(){
    if (__mmLockY === null) return;
    var b = document.body, y = __mmLockY; __mmLockY = null;
    b.style.position = ''; b.style.top = ''; b.style.left = ''; b.style.right = ''; b.style.width = '';
    window.scrollTo(0, y);
  }
  window.addEventListener('message', function(e){
    if (e.origin !== ORIGIN) return;            // origin 가드
    var d = e.data; if (!d) return;
    if (d.type === 'MAMORU_IFRAME_SIZE' && typeof d.height === 'number'){
      var list = frames();
      for (var i = 0; i < list.length; i++){
        if (list[i].contentWindow === e.source){  // source 매칭 = 보낸 iframe 정확히 지목
          var h = Math.ceil(d.height);
          var cur = parseInt(list[i].getAttribute('data-mm-h') || '0', 10);
          if (h > cur){                           // monotonic — 줄어들지 않음
            list[i].setAttribute('data-mm-h', String(h));
            list[i].style.height = h + 'px';
            list[i].style.minHeight = '0px';      // 혹시 남은 CSS 바닥값 해제(빈여백 방지)
          }
          return;
        }
      }
    } else if (d.type === 'MAMORU_NAVIGATE' && d.url){
      window.location.href = d.url;               // 자식 내부 링크 → 부모 네비게이션
    } else if (d.type === 'MAMORU_REQUEST_VIEWPORT'){
      // 자식 모달을 '현재 보이는 영역 중앙'에 띄우게 — 보낸 iframe의 가시 구간(visibleTop/Height, iframe 좌표) 회신
      var vl = frames();
      for (var v = 0; v < vl.length; v++){
        if (vl[v].contentWindow === e.source){
          var rect = vl[v].getBoundingClientRect();
          var vpH = window.innerHeight;
          var hdr = topFixedBottom();              // 상단 고정헤더(로고바) 아래
          var topEdge = Math.max(rect.top, hdr);   // 가시영역 상단 = 헤더 아래부터
          var visTop = Math.max(0, topEdge - rect.top);
          var visH = Math.max(0, Math.min(rect.bottom, vpH) - topEdge);
          try { e.source.postMessage({ type: 'MAMORU_VIEWPORT_INFO', visibleTop: visTop, visibleHeight: visH }, ORIGIN); } catch (_){}
          return;
        }
      }
    } else if (d.type === 'MAMORU_IFRAME_SCROLL' && typeof d.y === 'number'){
      // 자식이 '이 y 위치로 스크롤' 요청(임베드는 부모가 스크롤 주체) — 보낸 iframe의 문서top + y 로 이동
      var sl = frames();
      for (var s = 0; s < sl.length; s++){
        if (sl[s].contentWindow === e.source){
          var top = sl[s].getBoundingClientRect().top + (window.pageYOffset || 0);
          try { window.scrollTo({ top: Math.max(0, top + d.y - 16), behavior: 'smooth' }); }
          catch (_){ window.scrollTo(0, Math.max(0, top + d.y - 16)); }
          return;
        }
      }
    } else if (d.type === 'MAMORU_LOCK_SCROLL'){
      lockScroll();     // 모달 열림 — 배경 스크롤 잠금
    } else if (d.type === 'MAMORU_UNLOCK_SCROLL'){
      unlockScroll();   // 모달 닫힘 — 잠금 해제
    }
  });
  function requestAll(){
    frames().forEach(function(f){
      try { if (f.contentWindow) f.contentWindow.postMessage({ type: 'REQUEST_HEIGHT' }, ORIGIN); } catch (_){}
    });
  }
  // 로드~15초 넉넉히 재요청(자식 force 응답) + 늦게 주입되는 iframe 대비 20초간 주기 폴
  [300,800,1500,2500,4000,6000,9000,12000,15000].forEach(function(ms){ setTimeout(requestAll, ms); });
  var __mmReqTick = setInterval(requestAll, 1500);
  setTimeout(function(){ clearInterval(__mmReqTick); }, 20000);
  window.addEventListener('load', requestAll);
  if (document.readyState !== 'loading') requestAll();
})();
`;

export async function GET() {
  return new NextResponse(WIDGET_JS, {
    status: 200,
    headers: {
      'Content-Type': 'application/javascript; charset=utf-8',
      'Cache-Control': 'public, max-age=3600, s-maxage=3600',
      'Access-Control-Allow-Origin': '*',
    },
  });
}
