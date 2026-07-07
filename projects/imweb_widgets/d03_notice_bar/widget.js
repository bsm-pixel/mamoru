(function(){
  function initOne(root){
    if(root.getAttribute('data-flow')!=='true')return;
    var inEl=root.querySelector('.mm-notice__in');
    if(!inEl)return;
    // 끊김 없는 흐름: 내용 복제(50% 지점 동일)
    inEl.innerHTML=inEl.innerHTML+'<span style="display:inline-block;width:48px"></span>'+inEl.innerHTML;
  }
  function init(){var l=document.querySelectorAll('.mm-notice');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
