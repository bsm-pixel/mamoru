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

/* 동적 스타일(비율/배경/좌표): 인라인 style {{}}는 아임웹 저장거부 → data-* 속성을 JS로 적용 */
(function(){function ap(){
var A=document.querySelectorAll("[data-ar]");for(var i=0;i<A.length;i++){var v=(A[i].getAttribute("data-ar")||"").trim();if(v){if(A[i].classList.contains("mm-cat"))A[i].style.setProperty("--mm-ratio",v);else A[i].style.aspectRatio=v;}}
var B=document.querySelectorAll("[data-bg]");for(var i=0;i<B.length;i++){var v=(B[i].getAttribute("data-bg")||"").trim();if(v)B[i].style.backgroundImage="url('"+v+"')";}
var C=document.querySelectorAll("[data-x]");for(var i=0;i<C.length;i++){var x=(C[i].getAttribute("data-x")||"").trim(),y=(C[i].getAttribute("data-y")||"").trim();if(x)C[i].style.left=x+"%";if(y)C[i].style.top=y+"%";}
}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",ap);else ap();})();
