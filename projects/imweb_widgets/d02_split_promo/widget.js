/* 스플릿 프로모 배너 — 레이아웃은 HTML/CSS. JS는 높이만 적용(선택). 정규식·인라인핸들러 미사용. */
(function(){
  function initOne(root){
    var h=root.getAttribute('data-height');
    if(h){ var hv=h.trim(); if(hv){ if(String(parseFloat(hv))===hv) hv+='px'; root.style.minHeight=hv; } }
    /* 모바일 이미지가 있으면 data-hasm=1 → 모바일에서 스왑 */
    var m=root.querySelector('.mm-split__imgm');
    if(m && String(m.getAttribute('src')||'').trim()) root.setAttribute('data-hasm','1');
  }
  function init(){var l=document.querySelectorAll('.mm-split');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
