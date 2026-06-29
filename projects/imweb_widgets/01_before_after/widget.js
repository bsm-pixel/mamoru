(function(){
  function initOne(root){
    var stage=root.querySelector('.mm-ba__stage'),before=root.querySelector('.mm-ba__img--before'),handle=root.querySelector('.mm-ba__handle');
    if(!stage||!before||!handle)return;
    var start=parseFloat(root.getAttribute('data-start'));if(isNaN(start))start=50;
    var pos=Math.min(100,Math.max(0,start)),dragging=false;
    function apply(p){pos=Math.min(100,Math.max(0,p));var v='inset(0 '+(100-pos)+'% 0 0)';before.style.webkitClipPath=v;before.style.clipPath=v;handle.style.left=pos+'%';}
    function fromEvent(e){var r=stage.getBoundingClientRect();var cx=(e.touches&&e.touches[0])?e.touches[0].clientX:e.clientX;apply((cx-r.left)/r.width*100);}
    stage.addEventListener('pointerdown',function(e){dragging=true;try{stage.setPointerCapture(e.pointerId);}catch(_){}fromEvent(e);});
    stage.addEventListener('pointermove',function(e){if(dragging)fromEvent(e);});
    window.addEventListener('pointerup',function(){dragging=false;});
    root.setAttribute('tabindex','0');
    root.addEventListener('keydown',function(e){if(e.key==='ArrowLeft'){apply(pos-2);e.preventDefault();}else if(e.key==='ArrowRight'){apply(pos+2);e.preventDefault();}});
    apply(pos);
  }
  function init(){var l=document.querySelectorAll('.mm-ba');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
