(function(){
  function init(){
    var l=document.querySelectorAll('.mm-sh');
    if(!l.length){return setTimeout(init,50);}
    if(!('IntersectionObserver' in window)){for(var j=0;j<l.length;j++)l[j].classList.add('is-in');return;}
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('is-in');io.unobserve(e.target);}});},{threshold:.3});
    for(var i=0;i<l.length;i++)io.observe(l[i]);
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
