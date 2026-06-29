(function(){
  function initOne(root){
    var thumbs=root.querySelectorAll('.mm-gal__thumb');
    var lb=root.querySelector('.mm-gal__lb'),img=root.querySelector('.mm-gal__lbimg'),cap=root.querySelector('.mm-gal__cap');
    if(!thumbs.length||!lb)return;
    var cur=0;
    function open(i){cur=(i%thumbs.length+thumbs.length)%thumbs.length;img.src=thumbs[cur].getAttribute('data-img')||'';cap.textContent=thumbs[cur].getAttribute('data-cap')||'';lb.hidden=false;document.documentElement.style.overflow='hidden';}
    function close(){lb.hidden=true;document.documentElement.style.overflow='';}
    for(var x=0;x<thumbs.length;x++)(function(idx){thumbs[idx].addEventListener('click',function(){open(idx);});})(x);
    root.querySelector('.mm-gal__close').addEventListener('click',close);
    root.querySelector('.mm-gal__prev').addEventListener('click',function(){open(cur-1);});
    root.querySelector('.mm-gal__next').addEventListener('click',function(){open(cur+1);});
    lb.addEventListener('click',function(e){if(e.target===lb)close();});
    document.addEventListener('keydown',function(e){if(lb.hidden)return;if(e.key==='Escape')close();else if(e.key==='ArrowLeft')open(cur-1);else if(e.key==='ArrowRight')open(cur+1);});
  }
  function init(){var l=document.querySelectorAll('.mm-gal');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
