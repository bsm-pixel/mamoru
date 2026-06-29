(function(){
  function initOne(root){
    var spots=root.querySelectorAll('.mm-hs__spot');
    function closeAll(except){for(var k=0;k<spots.length;k++){if(spots[k]!==except)spots[k].classList.remove('is-open');}}
    for(var i=0;i<spots.length;i++)(function(sp){
      var dot=sp.querySelector('.mm-hs__dot');
      dot.addEventListener('click',function(e){e.stopPropagation();var on=sp.classList.contains('is-open');closeAll(sp);sp.classList.toggle('is-open',!on);});
    })(spots[i]);
    document.addEventListener('click',function(){closeAll(null);});
  }
  function init(){var l=document.querySelectorAll('.mm-hs');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
