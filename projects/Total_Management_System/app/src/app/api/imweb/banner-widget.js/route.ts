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
      '#mamoru-banner-close-x{position:absolute;top:10px;right:10px;width:32px;height:32px;border:none;background:rgba(250,249,247,0.9);border-radius:50%;font-size:18px;cursor:pointer;color:#1A1A1A;line-height:1;display:flex;align-items:center;justify-content:center;z-index:2;}',
      '#mamoru-banner-close-x:hover{background:#FFF;}',
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
        '<button id="mamoru-banner-close-x" aria-label="닫기">&times;</button>' +
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

    // X 버튼 + "닫기" 버튼
    overlay.querySelector('#mamoru-banner-close-x').addEventListener('click', close);
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
