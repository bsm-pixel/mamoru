(function(){
  function initOne(root){
    var stage=root.querySelector('.mm-ba__stage'),before=root.querySelector('.mm-ba__img--before'),handle=root.querySelector('.mm-ba__handle');
    if(!stage||!before||!handle)return;
    var start=parseFloat(root.getAttribute('data-start'));if(isNaN(start))start=50;
    var pos=Math.min(100,Math.max(0,start)),dragging=false;
    function apply(p){pos=Math.min(100,Math.max(0,p));var v='inset(0 '+(100-pos)+'% 0 0)';before.style.webkitClipPath=v;before.style.clipPath=v;handle.style.left=pos+'%';}
    function fromEvent(e){var r=stage.getBoundingClientRect();var cx=(e.touches&&e.touches[0])?e.touches[0].clientX:e.clientX;apply((cx-r.left)/r.width*100);}
    stage.addEventListener('pointerdown',function(e){dragging=true;try{stage.setPointerCapture(e.pointerId);}catch(_){}fromEvent(e);});
    stage.addEventListener('pointermove',function(e){if(dragging)fromEvent(e);});
    window.addEventListener('pointerup',function(){dragging=false;});
    root.setAttribute('tabindex','0');
    root.addEventListener('keydown',function(e){if(e.key==='ArrowLeft'){apply(pos-2);e.preventDefault();}else if(e.key==='ArrowRight'){apply(pos+2);e.preventDefault();}});
    apply(pos);
  }
  function init(){var l=document.querySelectorAll('.mm-ba');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();

/* 동적 스타일(비율/배경/좌표): 인라인 style {{}}는 아임웹 저장거부 → data-* 속성을 JS로 적용 */
(function(){function ap(){
var A=document.querySelectorAll("[data-ar]");for(var i=0;i<A.length;i++){var v=(A[i].getAttribute("data-ar")||"").trim();if(v){if(A[i].classList.contains("mm-cat"))A[i].style.setProperty("--mm-ratio",v);else A[i].style.aspectRatio=v;}}
var B=document.querySelectorAll("[data-bg]");for(var i=0;i<B.length;i++){var v=(B[i].getAttribute("data-bg")||"").trim();if(v)B[i].style.backgroundImage="url('"+v+"')";}
var C=document.querySelectorAll("[data-x]");for(var i=0;i<C.length;i++){var x=(C[i].getAttribute("data-x")||"").trim(),y=(C[i].getAttribute("data-y")||"").trim();if(x)C[i].style.left=x+"%";if(y)C[i].style.top=y+"%";}
}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",ap);else ap();})();
