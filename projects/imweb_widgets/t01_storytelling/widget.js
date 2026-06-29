(function(){
  function initOne(root){
    var media=root.querySelector('.mm-st__mediaimg');
    var scenes=root.querySelectorAll('.mm-st__scene');
    if(!scenes.length)return;
    if(media&&scenes[0])media.src=scenes[0].getAttribute('data-img')||'';
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){
        for(var k=0;k<scenes.length;k++)scenes[k].classList.toggle('is-active',scenes[k]===e.target);
        if(media){var img=e.target.getAttribute('data-img')||'';if(img&&media.src!==img){media.style.opacity='0';setTimeout(function(){media.src=img;media.style.opacity='1';},200);}}
      }});},{threshold:.5});
      for(var i=0;i<scenes.length;i++)io.observe(scenes[i]);
    }else{scenes[0].classList.add('is-active');}
  }
  function init(){var l=document.querySelectorAll('.mm-st');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
