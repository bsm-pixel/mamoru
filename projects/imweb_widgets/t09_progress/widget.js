(function(){
  function initOne(root){
    var bar=root.querySelector('.mm-prog__bar');
    var topBtn=root.querySelector('.mm-prog__top');
    var thick=parseFloat(root.getAttribute('data-thick'));
    if(bar&&!isNaN(thick)&&thick>0)bar.style.height=thick+'px';
    var useTop=root.getAttribute('data-backtop')!=='false';
    if(!useTop&&topBtn)topBtn.style.display='none';
    function onScroll(){
      var de=document.documentElement,b=document.body;
      var st=window.pageYOffset||de.scrollTop||b.scrollTop||0;
      var h=(de.scrollHeight||b.scrollHeight)-window.innerHeight;
      var p=h>0?Math.min(st/h,1):0;
      if(bar)bar.style.width=(p*100)+'%';
      if(useTop&&topBtn)topBtn.classList.toggle('is-show',st>window.innerHeight*0.6);
    }
    if(useTop&&topBtn)topBtn.addEventListener('click',function(){window.scrollTo({top:0,behavior:'smooth'});});
    window.addEventListener('scroll',onScroll,{passive:true});
    window.addEventListener('resize',onScroll);onScroll();
  }
  function init(){var l=document.querySelectorAll('.mm-prog');if(!l.length){return setTimeout(init,50);}initOne(l[0]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
