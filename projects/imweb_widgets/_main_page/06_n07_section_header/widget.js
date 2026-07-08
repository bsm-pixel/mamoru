(function(){
  /* 정렬(왼쪽/중앙/오른쪽)은 CSS 속성선택자가 담당(편집기 즉시반영). JS는 등장 애니만 */
  function init(){
    var l=document.querySelectorAll('.mm-sh');
    if(!l.length){return setTimeout(init,50);}
    function showAll(){for(var j=0;j<l.length;j++)l[j].classList.add('is-in');}
    if(!('IntersectionObserver' in window)){showAll();return;}
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('is-in');io.unobserve(e.target);}});},{threshold:.2});
    for(var i=0;i<l.length;i++)io.observe(l[i]);
    setTimeout(showAll,1400); /* 안전망: 아임웹 에디터 등 IntersectionObserver 미발동 시 강제 표시 */
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
