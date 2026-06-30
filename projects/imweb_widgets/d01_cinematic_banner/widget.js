/* 시네마틱 배너 — Ken Burns는 CSS. JS는 '어둡게' 색상 적용 + 정렬값 관대하게 인식. */
(function(){
  function initOne(root){
    // 어둡게: color-picker(검정+투명도) 값을 오버레이로
    var ov=root.getAttribute('data-overlay'),veil=root.querySelector('.mm-cine__veil');
    if(veil&&ov){veil.style.background=ov;}
    // 정렬: 가운데/왼쪽 외 유사어(중앙·center·좌측·left)도 인식
    var al=(root.getAttribute('data-align')||'').trim();
    root.setAttribute('data-align', /왼|좌|left|start/i.test(al) ? '왼쪽' : '가운데');
  }
  function init(){var l=document.querySelectorAll('.mm-cine');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
