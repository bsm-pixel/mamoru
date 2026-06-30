/* 스플릿 프로모 배너 — 순수 HTML/CSS 동작. JS 불필요(빈 안전망). */
(function(){
  function init(){var l=document.querySelectorAll('.mm-split');if(!l.length){return setTimeout(init,50);}}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
