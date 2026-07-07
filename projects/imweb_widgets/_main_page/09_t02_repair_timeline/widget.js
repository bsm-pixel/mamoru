(function(){
  function initOne(root){
    var steps=root.querySelectorAll('.mm-tl__step');
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('is-in');io.unobserve(e.target);}});},{threshold:.2,rootMargin:'0px 0px -10% 0px'});
      for(var i=0;i<steps.length;i++)io.observe(steps[i]);
    }else{for(var j=0;j<steps.length;j++)steps[j].classList.add('is-in');}
  }
  function init(){var l=document.querySelectorAll('.mm-tl');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
