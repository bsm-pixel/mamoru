/* 텍스트 마스크 — 순수 CSS(background-clip:text)로 동작. JS 불필요(진입 reveal만 선택). */
(function(){
  function init(){
    var l=document.querySelectorAll('.mm-mask');
    if(!l.length){return setTimeout(init,50);}
    if(!('IntersectionObserver' in window))return;
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.style.opacity='1';e.target.style.transform='none';io.unobserve(e.target);}});},{threshold:.2});
    for(var i=0;i<l.length;i++){l[i].style.transition='opacity .7s cubic-bezier(.4,0,.2,1),transform .7s cubic-bezier(.4,0,.2,1)';l[i].style.opacity='0';l[i].style.transform='translateY(20px)';io.observe(l[i]);}
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
