(function(){
  function fmt(n){return n.toLocaleString('en-US');}
  function run(root){
    var bigs=root.querySelectorAll('.mm-belief__big');
    for(var i=0;i<bigs.length;i++)(function(el){
      var raw=(el.textContent||'').trim();
      var suf=el.getAttribute('data-suffix')||'';
      if(/^[\d,]+$/.test(raw)){ // 숫자 → 카운트업
        var target=parseInt(raw.replace(/,/g,''),10)||0,t0=null,dur=1300;
        function step(ts){if(!t0)t0=ts;var p=Math.min((ts-t0)/dur,1),e=1-Math.pow(1-p,3);el.textContent=fmt(Math.round(target*e))+suf;if(p<1)requestAnimationFrame(step);}
        requestAnimationFrame(step);
      } else { el.textContent=raw+suf; } // 글자 → 그대로(+단위)
    })(bigs[i]);
  }
  function initOne(root){
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){root.classList.add('is-in');run(root);io.disconnect();}});},{threshold:.3});
      io.observe(root);
    }else{root.classList.add('is-in');run(root);}
  }
  function init(){var l=document.querySelectorAll('.mm-belief');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
