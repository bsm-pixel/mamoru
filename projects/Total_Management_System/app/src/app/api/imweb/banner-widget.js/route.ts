/**
 * GET /api/imweb/banner-widget.js
 * 아임웹 사이트에 주입되는 자기완결형 JavaScript
 *
 * 동작:
 *   1. DOMContentLoaded 후 /api/imweb/banner-config 호출
 *   2. enabled=true면 모달 DOM 생성
 *   3. "오늘 하루 보지 않기" 쿠키 + X 닫기 버튼
 *   4. MAMORU Brand Guide 준수 (모노크롬, Noto Sans KR)
 *
 * 아임웹 측 설치: <script src=".../api/imweb/banner-widget.js" defer></script>
 *
 * Content-Type: application/javascript
 * Cache-Control: 1시간 (widget 코드는 자주 안 바뀜)
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
      '#mamoru-banner-image{width:100%;display:block;background:#EEE;}',
      '#mamoru-banner-image.clickable{cursor:pointer;}',
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

  function render(config) {
    injectStyles();

    var overlay = document.createElement('div');
    overlay.id = 'mamoru-banner-overlay';

    var imgHtml = '';
    if (config.image_url) {
      var cls = config.link_url ? 'clickable' : '';
      imgHtml = '<img id="mamoru-banner-image" class="' + cls + '" src="' + esc(config.image_url) + '" alt="' + esc(config.title || '배너') + '">';
    }

    var descHtml = config.description
      ? '<p id="mamoru-banner-desc">' + esc(config.description) + '</p>'
      : '';

    overlay.innerHTML =
      '<div id="mamoru-banner-modal" role="dialog" aria-modal="true">' +
        '<button id="mamoru-banner-close-x" aria-label="닫기">&times;</button>' +
        imgHtml +
        (config.title || config.description
          ? '<div id="mamoru-banner-body">' +
              (config.title ? '<h3 id="mamoru-banner-title">' + esc(config.title) + '</h3>' : '') +
              descHtml +
            '</div>'
          : '') +
        '<div id="mamoru-banner-actions">' +
          '<button id="mamoru-banner-dismiss">오늘 하루 보지 않기</button>' +
          '<button id="mamoru-banner-close" class="primary">닫기</button>' +
        '</div>' +
      '</div>';

    function close() {
      if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
    }

    // X 버튼 + "닫기" 버튼: 단순 닫기
    overlay.querySelector('#mamoru-banner-close-x').addEventListener('click', close);
    overlay.querySelector('#mamoru-banner-close').addEventListener('click', close);

    // "오늘 하루 보지 않기": 쿠키 저장 + 닫기
    overlay.querySelector('#mamoru-banner-dismiss').addEventListener('click', function() {
      setCookie(COOKIE_KEY, Date.now(), config.dismiss_cookie_hours || 24);
      close();
    });

    // 이미지 클릭 → link_url 이동
    if (config.link_url) {
      var img = overlay.querySelector('#mamoru-banner-image');
      if (img) {
        img.addEventListener('click', function() {
          window.open(config.link_url, '_blank', 'noopener');
        });
      }
    }

    // 배경 클릭 시 닫기 (모달 내부 클릭은 전파 차단)
    overlay.addEventListener('click', function(e) {
      if (e.target === overlay) close();
    });

    document.body.appendChild(overlay);
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
