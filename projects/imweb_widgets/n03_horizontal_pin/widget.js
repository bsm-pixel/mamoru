(function(){
  function initOne(root){
    var sticky=root.querySelector('.mm-hpin__sticky');
    var track=root.querySelector('.mm-hpin__track');
    if(!sticky||!track)return;
    var reduce=window.matchMedia&&window.matchMedia('(prefers-reduced-motion:reduce)').matches;
    var touch=window.matchMedia&&window.matchMedia('(hover:none)').matches;
    function setup(){
      var dist=track.scrollWidth-window.innerWidth;
      if(reduce||touch||dist<=40){root.classList.add('is-native');root.style.height='';track.style.transform='';return;}
      root.classList.remove('is-native');
      root.style.height=(dist+window.innerHeight)+'px';
      onScroll();
    }
    function onScroll(){
      if(root.classList.contains('is-native'))return;
      var rect=root.getBoundingClientRect();
      var dist=track.scrollWidth-window.innerWidth;
      var p=Math.min(Math.max(-rect.top/(root.offsetHeight-window.innerHeight),0),1);
      track.style.transform='translateX('+(-p*dist)+'px)';
    }
    window.addEventListener('scroll',onScroll,{passive:true});
    window.addEventListener('resize',setup);
    setup();
    // 이미지 로드 후 폭 재계산
    var imgs=root.querySelectorAll('img');for(var i=0;i<imgs.length;i++)imgs[i].addEventListener('load',setup);
  }
  function init(){var l=document.querySelectorAll('.mm-hpin');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
