(function(){
  function initOne(root){
    var h=root.getAttribute('data-height');
    if(h){ var hv=h.trim(); if(hv){ if(String(parseFloat(hv))===hv) hv+='px'; root.style.minHeight=hv; } }
    root.addEventListener('pointermove',function(e){
      if(e.pointerType==='touch')return;
      var r=root.getBoundingClientRect();
      root.style.setProperty('--mx',((e.clientX-r.left)/r.width*100)+'%');
      root.style.setProperty('--my',((e.clientY-r.top)/r.height*100)+'%');
    });
  }
  function init(){var l=document.querySelectorAll('.mm-sp');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
