(function(){
  function bind(btn){
    var STR=0.35,MAX=14;
    btn.addEventListener('pointermove',function(e){
      if(e.pointerType==='touch')return;
      var r=btn.getBoundingClientRect();
      var x=(e.clientX-(r.left+r.width/2))*STR;
      var y=(e.clientY-(r.top+r.height/2))*STR;
      x=Math.max(-MAX,Math.min(MAX,x));y=Math.max(-MAX,Math.min(MAX,y));
      btn.style.transform='translate('+x+'px,'+y+'px)';
    });
    btn.addEventListener('pointerleave',function(){btn.style.transform='';});
  }
  function init(){var l=document.querySelectorAll('.mm-mag__btn');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)bind(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
