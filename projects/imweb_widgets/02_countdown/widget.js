(function(){
  function parseTarget(dateStr,timeStr){
    if(!dateStr) return null;
    var dm=String(dateStr).match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
    if(!dm) return null;
    var y=+dm[1],mo=+dm[2]-1,d=+dm[3],h=23,mi=59;
    if(timeStr){
      var pm=/오후|PM/i.test(timeStr),am=/오전|AM/i.test(timeStr);
      var tm=String(timeStr).match(/(\d{1,2})\D+(\d{2})/);
      if(tm){h=+tm[1];mi=+tm[2];if(pm&&h<12)h+=12;if(am&&h===12)h=0;}
    }
    var t=new Date(y,mo,d,h,mi,0).getTime();
    return isNaN(t)?null:t;
  }
  function pad(n){return n<10?'0'+n:''+n;}
  function initOne(root){
    var target=parseTarget(root.getAttribute('data-date'),root.getAttribute('data-time'));
    var endedEl=root.querySelector('.mm-cd__ended');
    var endedText=endedEl?endedEl.getAttribute('data-ended'):'';
    if(endedEl&&endedText) endedEl.textContent=endedText;
    var nums={d:root.querySelector('[data-cd="d"]'),h:root.querySelector('[data-cd="h"]'),m:root.querySelector('[data-cd="m"]'),s:root.querySelector('[data-cd="s"]')};
    if(!target){root.classList.add('is-ended');return;}
    var timer=null;
    function tick(){
      var diff=target-Date.now();
      if(diff<=0){root.classList.add('is-ended');if(timer)clearInterval(timer);return;}
      var s=Math.floor(diff/1000);
      var d=Math.floor(s/86400);s-=d*86400;
      var h=Math.floor(s/3600);s-=h*3600;
      var m=Math.floor(s/60);s-=m*60;
      if(nums.d)nums.d.textContent=pad(d);
      if(nums.h)nums.h.textContent=pad(h);
      if(nums.m)nums.m.textContent=pad(m);
      if(nums.s)nums.s.textContent=pad(s);
    }
    tick();timer=setInterval(tick,1000);
  }
  function init(){var l=document.querySelectorAll('.mm-cd');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
