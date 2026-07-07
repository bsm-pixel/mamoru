(function(){
  /* 숫자값 → 단위 부여(폰트 pt / 여백 px) 후 CSS 변수 주입 */
  function unit(v,u){v=String(v==null?'':v).trim(); if(!v)return null; if(String(parseFloat(v))===v)v+=u; return v;}
  function applySettings(root){
    var fpc=unit(root.getAttribute('data-fpc'),'pt'); if(fpc)root.style.setProperty('--nav-fpc',fpc);
    var fm=unit(root.getAttribute('data-fm'),'pt'); if(fm)root.style.setProperty('--nav-fm',fm);
    var py=unit(root.getAttribute('data-pady'),'px'); if(py)root.style.setProperty('--nav-pady',py);
  }
  /* 업로드한 아이콘 이미지(SVG/PNG)를 mask로 적용 → background(currentColor)가 테마색으로 자동 채색 */
  function applyIcons(root){
    var icos=root.querySelectorAll('.mm-nav__ico');
    for(var i=0;i<icos.length;i++){
      var src=String(icos[i].getAttribute('data-src')||'').trim();
      if(src){ var u="url('"+src+"')"; icos[i].style.webkitMaskImage=u; icos[i].style.maskImage=u; }
    }
  }
  function initOne(root){
    applySettings(root);
    applyIcons(root);
    /* 폰트·여백 값 바뀌면 즉시 반영(편집기 실시간). 정렬(가운데/왼쪽)은 CSS 속성선택자가 담당 */
    if('MutationObserver' in window){ new MutationObserver(function(){applySettings(root);}).observe(root,{attributes:true,attributeFilter:['data-fpc','data-fm','data-pady']}); }
  }
  function init(){var l=document.querySelectorAll('.mm-nav');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
