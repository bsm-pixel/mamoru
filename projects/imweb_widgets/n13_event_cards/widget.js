(function(){
  /* 가로 최대폭: 자유 숫자값(1280 등). 비움/full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  function applySettings(root){
    var c=String(root.getAttribute('data-cols')||'').trim();
    if(/^[1-9]\d*$/.test(c)) root.style.setProperty('--ev-cols',c);
    applyMaxw(root);
  }
  function initOne(root){
    applySettings(root);
    /* PC 열수·최대폭 바뀌면 즉시 반영(편집기 실시간) */
    if('MutationObserver' in window){ new MutationObserver(function(){applySettings(root);}).observe(root,{attributes:true,attributeFilter:['data-cols','data-maxw']}); }
    /* 카드 순차 등장 */
    var cards=root.querySelectorAll('.mm-ev__card');
    function reveal(){ for(var k=0;k<cards.length;k++){ cards[k].style.transitionDelay=((k%12)*0.06)+'s'; cards[k].classList.add('is-in'); } }
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){reveal();io.disconnect();}});},{threshold:.1});
      io.observe(root);
    }else{ reveal(); }
  }
  function init(){var l=document.querySelectorAll('.mm-ev');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
