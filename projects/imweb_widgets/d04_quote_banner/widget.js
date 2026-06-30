/* 대형 인용 배너 — 순수 CSS 동작. 진입 reveal(선택). */
(function(){
  function init(){
    var l=document.querySelectorAll('.mm-quote');
    if(!l.length){return setTimeout(init,50);}
    if(!('IntersectionObserver' in window))return;
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.style.opacity='1';e.target.style.transform='none';io.unobserve(e.target);}});},{threshold:.3});
    for(var i=0;i<l.length;i++){l[i].style.transition='opacity .7s cubic-bezier(.4,0,.2,1),transform .7s cubic-bezier(.4,0,.2,1)';l[i].style.opacity='0';l[i].style.transform='translateY(18px)';io.observe(l[i]);}
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
