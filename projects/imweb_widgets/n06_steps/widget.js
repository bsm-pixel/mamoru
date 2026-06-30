/* 이용 안내 가로 스텝 — CSS 카운터로 번호 자동. 진입 reveal(선택). */
(function(){
  function init(){
    var l=document.querySelectorAll('.mm-step');
    if(!l.length){return setTimeout(init,50);}
    if(!('IntersectionObserver' in window))return;
    for(var w=0;w<l.length;w++){
      var items=l[w].querySelectorAll('.mm-step__item');
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.style.opacity='1';e.target.style.transform='none';io.unobserve(e.target);}});},{threshold:.25});
      for(var i=0;i<items.length;i++){items[i].style.transition='opacity .5s cubic-bezier(.4,0,.2,1) '+(i*0.08)+'s, transform .5s cubic-bezier(.4,0,.2,1) '+(i*0.08)+'s';items[i].style.opacity='0';items[i].style.transform='translateY(14px)';io.observe(items[i]);}
    }
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
