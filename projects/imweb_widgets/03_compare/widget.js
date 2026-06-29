/* 스펙 비교표 — 카드 진입 reveal(선택). 핵심 기능은 HTML/CSS만으로 동작. */
(function(){
  function initOne(root){
    if(!('IntersectionObserver' in window))return;
    var cards=root.querySelectorAll('.mm-cmp__card');
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.style.opacity='1';e.target.style.transform='none';io.unobserve(e.target);}});},{threshold:.15});
    for(var i=0;i<cards.length;i++){cards[i].style.transition='opacity .5s cubic-bezier(.4,0,.2,1) '+(i*0.06)+'s, transform .5s cubic-bezier(.4,0,.2,1) '+(i*0.06)+'s';cards[i].style.opacity='0';cards[i].style.transform='translateY(18px)';io.observe(cards[i]);}
  }
  function init(){var l=document.querySelectorAll('.mm-cmp');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
