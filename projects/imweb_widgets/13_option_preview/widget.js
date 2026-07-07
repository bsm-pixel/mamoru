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

/* 동적 스타일(비율/배경/좌표): 인라인 style {{}}는 아임웹 저장거부 → data-* 속성을 JS로 적용 */
(function(){function ap(){
var A=document.querySelectorAll("[data-ar]");for(var i=0;i<A.length;i++){var v=(A[i].getAttribute("data-ar")||"").trim();if(v){if(A[i].classList.contains("mm-cat"))A[i].style.setProperty("--mm-ratio",v);else A[i].style.aspectRatio=v;}}
var B=document.querySelectorAll("[data-bg]");for(var i=0;i<B.length;i++){var v=(B[i].getAttribute("data-bg")||"").trim();if(v)B[i].style.backgroundImage="url('"+v+"')";}
var C=document.querySelectorAll("[data-x]");for(var i=0;i<C.length;i++){var x=(C[i].getAttribute("data-x")||"").trim(),y=(C[i].getAttribute("data-y")||"").trim();if(x)C[i].style.left=x+"%";if(y)C[i].style.top=y+"%";}
}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",ap);else ap();})();
