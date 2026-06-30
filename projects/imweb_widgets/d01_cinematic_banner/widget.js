/* 시네마틱 배너 — Ken Burns는 CSS. '어둡게'=검정 고정+투명도만, 정렬 관대 인식. (정규식 미사용) */
(function(){
  function alphaOf(ov){
    ov=String(ov||'').trim();
    if(ov.indexOf('rgba')===0){
      var inner=ov.substring(ov.indexOf('(')+1, ov.lastIndexOf(')'));
      var p=inner.split(',');
      if(p.length>=4){ var a=parseFloat(p[3]); if(!isNaN(a)) return a; }
      return null;
    }
    var hex=ov.charAt(0)==='#'?ov.substring(1):ov;
    if(hex.length===8){ var v=parseInt(hex.substring(6,8),16); if(!isNaN(v)) return v/255; }
    return null;
  }
  function isLeft(al){
    al=String(al||'').trim(); var low=al.toLowerCase();
    return al.indexOf('왼')>=0 || al.indexOf('좌')>=0 || low.indexOf('left')>=0 || low.indexOf('start')>=0;
  }
  function initOne(root){
    var ov=root.getAttribute('data-overlay'),veil=root.querySelector('.mm-cine__veil');
    if(veil&&ov){ var a=alphaOf(ov); if(a!==null){ a=Math.max(0,Math.min(1,a)); veil.style.background='rgba(26,26,26,'+a+')'; } else { veil.style.background=ov; } }
    root.setAttribute('data-align', isLeft(root.getAttribute('data-align')) ? '왼쪽' : '가운데');
  }
  function init(){var l=document.querySelectorAll('.mm-cine');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
