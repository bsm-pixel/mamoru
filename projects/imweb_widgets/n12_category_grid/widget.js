(function(){
  function unit(v,u){v=String(v==null?'':v).trim(); if(!v)return null; if(String(parseFloat(v))===v)v+=u; return v;}
  /* 가로 최대폭: 자유 숫자값(1280 등). 비움/full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  /* PC/모바일 타일 높이 → CSS 변수 주입 */
  function applySettings(root){
    var hpc=unit(root.getAttribute('data-hpc'),'px'); if(hpc)root.style.setProperty('--cat-h',hpc);
    var hm=unit(root.getAttribute('data-hm'),'px'); if(hm)root.style.setProperty('--cat-hm',hm);
    applyMaxw(root);
  }
  function initOne(root){
    applySettings(root);
    /* 타일별 모바일 이미지가 실제 있으면 data-hasm=1 → 모바일에서 스왑 */
    var tiles=root.querySelectorAll('.mm-cat__tile');
    for(var i=0;i<tiles.length;i++){
      var m=tiles[i].querySelector('.mm-cat__imgm');
      if(m && String(m.getAttribute('src')||'').trim()) tiles[i].setAttribute('data-hasm','1');
    }
    /* 높이·최대폭 값 바뀌면 즉시 반영(편집기 실시간). 정렬은 CSS 속성선택자가 담당 */
    if('MutationObserver' in window){ new MutationObserver(function(){applySettings(root);}).observe(root,{attributes:true,attributeFilter:['data-hpc','data-hm','data-maxw']}); }
  }
  function init(){var l=document.querySelectorAll('.mm-cat');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
