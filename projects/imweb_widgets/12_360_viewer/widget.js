(function(){
  function initOne(root){
    var stage=root.querySelector('.mm-360__stage');
    var frames=root.querySelectorAll('.mm-360__f');
    var F=frames.length;
    if(!stage||!F)return;
    var idx=0,dragging=false,startX=0,startIdx=0;
    function show(i){idx=((i%F)+F)%F;for(var k=0;k<F;k++)frames[k].classList.toggle('is-on',k===idx);}
    function move(x){var step=stage.getBoundingClientRect().width/F||1;var d=Math.round((x-startX)/step);show(startIdx-d);}
    stage.addEventListener('pointerdown',function(e){dragging=true;root.classList.add('is-active');startX=e.clientX;startIdx=idx;try{stage.setPointerCapture(e.pointerId);}catch(_){}});
    stage.addEventListener('pointermove',function(e){if(dragging)move(e.clientX);});
    window.addEventListener('pointerup',function(){dragging=false;});
    show(0);
  }
  function init(){var l=document.querySelectorAll('.mm-360');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
