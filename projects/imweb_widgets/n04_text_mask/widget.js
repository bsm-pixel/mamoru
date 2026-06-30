/* 텍스트 마스크 — 정렬값 관대 인식 + 순수 CSS(background-clip:text). 진입 reveal. */
(function(){
  function isLeft(al){al=String(al||'').trim();var low=al.toLowerCase();return al==='true'||al.indexOf('왼')>=0||al.indexOf('좌')>=0||low.indexOf('left')>=0||low.indexOf('start')>=0;}
  function init(){
    var l=document.querySelectorAll('.mm-mask');
    if(!l.length){return setTimeout(init,50);}
    for(var i=0;i<l.length;i++) l[i].setAttribute('data-align', isLeft(l[i].getAttribute('data-align'))?'왼쪽':'가운데');
    if(!('IntersectionObserver' in window))return;
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.style.opacity='1';e.target.style.transform='none';io.unobserve(e.target);}});},{threshold:.2});
    for(var j=0;j<l.length;j++){l[j].style.transition='opacity .7s cubic-bezier(.4,0,.2,1),transform .7s cubic-bezier(.4,0,.2,1)';l[j].style.opacity='0';l[j].style.transform='translateY(20px)';io.observe(l[j]);}
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
