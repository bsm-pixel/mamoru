/* 시네마틱 배너 — Ken Burns는 CSS. JS는 '어둡게' 색상값을 오버레이에 적용. */
(function(){
  function initOne(root){
    var ov=root.getAttribute('data-overlay');
    var veil=root.querySelector('.mm-cine__veil');
    if(veil&&ov){veil.style.background=ov;} // color-picker 값(검정+투명도)을 그대로 오버레이로
  }
  function init(){var l=document.querySelectorAll('.mm-cine');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
