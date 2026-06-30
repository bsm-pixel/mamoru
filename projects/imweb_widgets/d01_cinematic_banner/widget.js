/* 시네마틱 배너 — Ken Burns는 CSS. JS: '어둡게'는 색 무시·보이드 검정 고정 + 투명도만 반영(브랜드 일관). 정렬 관대 인식. */
(function(){
  function applyVeil(veil,ov){
    var a=null,m;
    m=String(ov).match(/rgba\([^)]*,\s*([\d.]+)\s*\)/); if(m)a=parseFloat(m[1]);              // rgba(...,0.27)
    else { m=String(ov).match(/^#?[0-9a-fA-F]{6}([0-9a-fA-F]{2})$/); if(m)a=parseInt(m[1],16)/255; } // #RRGGBBAA
    if(a!=null){ a=Math.max(0,Math.min(1,a)); veil.style.background='rgba(26,26,26,'+a+')'; } // 투명도만 → 검정 고정
    else veil.style.background=ov;                                                            // 형식 모르면 picked 값 그대로(안전)
  }
  function initOne(root){
    var ov=root.getAttribute('data-overlay'),veil=root.querySelector('.mm-cine__veil');
    if(veil&&ov)applyVeil(veil,ov);
    var al=(root.getAttribute('data-align')||'').trim();
    root.setAttribute('data-align', /왼|좌|left|start/i.test(al) ? '왼쪽' : '가운데');
  }
  function init(){var l=document.querySelectorAll('.mm-cine');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
