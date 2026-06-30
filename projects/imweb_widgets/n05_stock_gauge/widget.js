(function(){
  function num(v,d){var n=parseFloat(String(v).replace(/[^0-9.]/g,''));return isNaN(n)?d:n;}
  function initOne(root){
    var total=num(root.getAttribute('data-total'),100);
    var rem=num(root.getAttribute('data-remaining'),0);
    var unit=root.getAttribute('data-unit')||'';
    var bar=root.querySelector('.mm-stock__bar');
    var count=root.querySelector('.mm-stock__count');
    rem=Math.max(0,Math.min(rem,total));
    var pct=total>0?rem/total*100:0;
    if(count)count.innerHTML='<b>'+rem+'</b>'+unit+' 남음';
    root.classList.toggle('is-low',pct<=30);
    function go(){if(bar)bar.style.width=pct+'%';}
    if('IntersectionObserver' in window){var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){go();io.disconnect();}});},{threshold:.4});io.observe(root);}else{go();}
  }
  function init(){var l=document.querySelectorAll('.mm-stock');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
