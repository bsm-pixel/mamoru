(function(){
  function bind(card){
    var inner=card.querySelector('.mm-tilt__inner');
    if(!inner)return;
    var MAX=8;
    card.addEventListener('pointermove',function(e){
      if(e.pointerType==='touch')return;
      var r=card.getBoundingClientRect();
      var px=(e.clientX-r.left)/r.width-.5, py=(e.clientY-r.top)/r.height-.5;
      inner.style.transform='rotateY('+(px*MAX)+'deg) rotateX('+(-py*MAX)+'deg)';
    });
    card.addEventListener('pointerleave',function(){inner.style.transform='';});
  }
  function initOne(root){var cards=root.querySelectorAll('.mm-tilt__card');for(var i=0;i<cards.length;i++)bind(cards[i]);}
  function init(){var l=document.querySelectorAll('.mm-tilt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
