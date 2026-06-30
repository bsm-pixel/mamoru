/* 시네마틱 배너 — 순수 CSS(Ken Burns) 동작. JS 불필요(빈 안전망만). */
(function(){
  function init(){var l=document.querySelectorAll('.mm-cine');if(!l.length){return setTimeout(init,50);}}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
