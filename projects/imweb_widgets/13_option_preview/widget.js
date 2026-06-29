(function(){
  function initOne(root){
    var main=root.querySelector('.mm-opt__main');
    var btns=root.querySelectorAll('.mm-opt__btn');
    if(!main||!btns.length)return;
    function swap(i){
      var img=btns[i].getAttribute('data-img')||'';
      main.style.opacity='0';
      setTimeout(function(){main.src=img;main.style.opacity='1';},150);
      for(var k=0;k<btns.length;k++)btns[k].classList.toggle('is-on',k===i);
    }
    for(var x=0;x<btns.length;x++)(function(idx){btns[idx].addEventListener('click',function(){swap(idx);});})(x);
    // 초기: 첫 옵션
    main.src=btns[0].getAttribute('data-img')||'';btns[0].classList.add('is-on');
  }
  function init(){var l=document.querySelectorAll('.mm-opt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
