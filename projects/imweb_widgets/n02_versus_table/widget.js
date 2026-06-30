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
