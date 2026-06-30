/* 카테고리 이미지 그리드 — 순수 HTML/CSS 동작. 진입 reveal(선택). */
(function(){
  function initOne(root){
    if(!('IntersectionObserver' in window))return;
    var tiles=root.querySelectorAll('.mm-cat__tile');
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.style.opacity='1';e.target.style.transform='none';io.unobserve(e.target);}});},{threshold:.12});
    for(var i=0;i<tiles.length;i++){tiles[i].style.transition='opacity .5s cubic-bezier(.4,0,.2,1) '+(i*0.05)+'s, transform .5s cubic-bezier(.4,0,.2,1) '+(i*0.05)+'s';tiles[i].style.opacity='0';tiles[i].style.transform='translateY(16px)';io.observe(tiles[i]);}
  }
  function init(){var l=document.querySelectorAll('.mm-cat');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
