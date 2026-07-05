/* 시네마틱 배너 — Ken Burns는 CSS. '어둡게'=검정+투명도, 정렬·사진초점·높이 적용. (정규식·인라인핸들러 미사용) */
(function(){
  function alphaOf(ov){
    ov=String(ov||'').trim();
    if(ov.indexOf('rgba')===0){
      var inner=ov.substring(ov.indexOf('(')+1, ov.lastIndexOf(')')),p=inner.split(',');
      if(p.length>=4){ var a=parseFloat(p[3]); if(!isNaN(a)) return a; }
      return null;
    }
    var hex=ov.charAt(0)==='#'?ov.substring(1):ov;
    if(hex.length===8){ var v=parseInt(hex.substring(6,8),16); if(!isNaN(v)) return v/255; }
    return null;
  }
  function isLeft(al){al=String(al||'').trim();var low=al.toLowerCase();
    /* 어떤 유형이 와도 인식: 텍스트(왼쪽/좌/left) · 스위치(true) · 옵션버튼(왼쪽) */
    return al==='true'||al.indexOf('왼')>=0||al.indexOf('좌')>=0||low.indexOf('left')>=0||low.indexOf('start')>=0;}
  function focusPos(f){
    f=String(f||'').trim();var low=f.toLowerCase();
    if(f.indexOf('좌')>=0||f.indexOf('왼')>=0||low.indexOf('left')>=0)return 'left center';
    if(f.indexOf('우')>=0||f.indexOf('오')>=0||low.indexOf('right')>=0)return 'right center';
    if(f.indexOf('상')>=0||f.indexOf('위')>=0||low.indexOf('top')>=0)return 'center top';
    if(f.indexOf('하')>=0||f.indexOf('아래')>=0||low.indexOf('bottom')>=0)return 'center bottom';
    return 'center';
  }
  function initOne(root){
    var ov=root.getAttribute('data-overlay'),veil=root.querySelector('.mm-cine__veil'),bg=root.querySelector('.mm-cine__bg');
    if(veil&&ov){ var a=alphaOf(ov); if(a!==null){ a=Math.max(0,Math.min(1,a)); veil.style.background='rgba(26,26,26,'+a+')'; } else { veil.style.background=ov; } }
    root.setAttribute('data-align', isLeft(root.getAttribute('data-align')) ? '왼쪽' : '가운데');
    if(bg) bg.style.objectPosition=focusPos(root.getAttribute('data-focus'));
    var h=root.getAttribute('data-height');
    if(h){ var hv=h.trim(); if(hv){ if(String(parseFloat(hv))===hv) hv+='px'; root.style.minHeight=hv; } }
    var r=root.getAttribute('data-radius');
    if(r!==null){ var rv=r.trim(); if(rv){ if(String(parseFloat(rv))===rv) rv+='px'; root.style.borderRadius=rv; } }
  }
  function init(){var l=document.querySelectorAll('.mm-cine');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
