(function(){
  function initOne(root){
    var stage=root.querySelector('.mm-360__stage');
    var frames=root.querySelectorAll('.mm-360__f');
    var F=frames.length;
    if(!stage||!F)return;
    var idx=0,dragging=false,startX=0,startIdx=0;
    function show(i){idx=((i%F)+F)%F;for(var k=0;k<F;k++)frames[k].classList.toggle('is-on',k===idx);}
    function move(x){var step=stage.getBoundingClientRect().width/F||1;var d=Math.round((x-startX)/step);show(startIdx-d);}
    stage.addEventListener('pointerdown',function(e){dragging=true;root.classList.add('is-active');startX=e.clientX;startIdx=idx;try{stage.setPointerCapture(e.pointerId);}catch(_){}});
    stage.addEventListener('pointermove',function(e){if(dragging)move(e.clientX);});
    window.addEventListener('pointerup',function(){dragging=false;});
    show(0);
  }
  function init(){var l=document.querySelectorAll('.mm-360');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 동적 스타일(비율/배경/좌표): 인라인 style {{}}는 아임웹 저장거부 → data-* 속성을 JS로 적용 */
(function(){function ap(){
var A=document.querySelectorAll("[data-ar]");for(var i=0;i<A.length;i++){var v=(A[i].getAttribute("data-ar")||"").trim();if(v){if(A[i].classList.contains("mm-cat"))A[i].style.setProperty("--mm-ratio",v);else A[i].style.aspectRatio=v;}}
var B=document.querySelectorAll("[data-bg]");for(var i=0;i<B.length;i++){var v=(B[i].getAttribute("data-bg")||"").trim();if(v)B[i].style.backgroundImage="url('"+v+"')";}
var C=document.querySelectorAll("[data-x]");for(var i=0;i<C.length;i++){var x=(C[i].getAttribute("data-x")||"").trim(),y=(C[i].getAttribute("data-y")||"").trim();if(x)C[i].style.left=x+"%";if(y)C[i].style.top=y+"%";}
}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",ap);else ap();})();
