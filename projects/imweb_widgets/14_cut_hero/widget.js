(function(){
  function initOne(root){
    var m=root.querySelector('.mm-cut__bgm');
    if(m && String(m.getAttribute('src')||'').trim()) root.setAttribute('data-hasm','1');
    /* 높이: 비우면 이미지 비율 자동(CSS). 값 입력 시 고정 높이+크롭(data-fixed) → 텍스트가 안 밀어 정확 */
    var h=root.getAttribute('data-height');
    if(h && h.trim()){ var hv=h.trim(); if(String(parseFloat(hv))===hv) hv+='px'; root.style.minHeight=hv; root.setAttribute('data-fixed','1'); }
    else {
      root.removeAttribute('data-fixed');
      /* 이미지 있으면 바닥값 없이 이미지 비율 그대로(짧으면 짧게). 이미지 없을 때만 바닥 높이 */
      var bg=root.querySelector('.mm-cut__bg');
      var hasImg=(bg&&String(bg.getAttribute('src')||'').trim())||(m&&String(m.getAttribute('src')||'').trim());
      root.style.minHeight = hasImg ? '' : 'clamp(220px,42vw,440px)';
    }
    /* 배너 실제 높이 → --cut-h 주입 → CSS가 폰트·간격을 높이 비례로 스케일(낮은 배너서 안 눌림). 이미지로드·리사이즈·높이변경 자동 반영 */
    function applyScale(){ var hh=Math.round(root.getBoundingClientRect().height); if(hh>0) root.style.setProperty('--cut-h', hh); }
    applyScale();
    if('ResizeObserver' in window){ new ResizeObserver(applyScale).observe(root); }
    else { window.addEventListener('resize', applyScale); }
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('is-in');io.unobserve(e.target);}});},{threshold:.3});
      io.observe(root);
    }else{root.classList.add('is-in');}
  }
  function init(){var l=document.querySelectorAll('.mm-cut');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
