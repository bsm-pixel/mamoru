(function(){
  function isLeft(al){al=String(al||'').trim();var low=al.toLowerCase();return al==='true'||al.indexOf('왼')>=0||al.indexOf('좌')>=0||low.indexOf('left')>=0||low.indexOf('start')>=0;}
  function init(){
    var l=document.querySelectorAll('.mm-sh');
    if(!l.length){return setTimeout(init,50);}
    for(var k=0;k<l.length;k++) l[k].setAttribute('data-align', isLeft(l[k].getAttribute('data-align'))?'왼쪽':'가운데');
    if(!('IntersectionObserver' in window)){for(var j=0;j<l.length;j++)l[j].classList.add('is-in');return;}
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('is-in');io.unobserve(e.target);}});},{threshold:.3});
    for(var i=0;i<l.length;i++)io.observe(l[i]);
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
