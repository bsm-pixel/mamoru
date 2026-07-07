(function(){
  function render(cell){
    var v=(cell.getAttribute('data-v')||'').trim();
    var low=v.toLowerCase();
    if(low==='o'||low==='yes'||low==='y'||v==='✓'||v==='○'){cell.classList.add('is-o');cell.textContent='';}
    else if(low==='x'||low==='no'||low==='n'||v==='✕'||v==='×'){cell.classList.add('is-x');cell.textContent='';}
    else{cell.textContent=v;}
  }
  function initOne(root){var cells=root.querySelectorAll('.mm-vs__cell');for(var i=0;i<cells.length;i++)render(cells[i]);}
  function init(){var l=document.querySelectorAll('.mm-vs');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
