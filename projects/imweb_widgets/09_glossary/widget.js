(function(){
  function initOne(root){
    var search=root.querySelector('.mm-gl__search');
    var items=root.querySelectorAll('.mm-gl__item');
    var empty=root.querySelector('.mm-gl__empty');
    if(!search)return;
    function txt(el){return (el.textContent||'').toLowerCase();}
    search.addEventListener('input',function(){
      var q=search.value.trim().toLowerCase(),shown=0;
      for(var i=0;i<items.length;i++){var hit=!q||txt(items[i]).indexOf(q)>=0;items[i].classList.toggle('is-hide',!hit);if(hit)shown++;}
      if(empty)empty.hidden=shown>0;
    });
  }
  function init(){var l=document.querySelectorAll('.mm-gl');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
