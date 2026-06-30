(function(){
  function num(v,d){var n=parseFloat(v);return isNaN(n)?d:n;}
  function initOne(root){
    var range=root.querySelector('.mm-rul__range');
    var bar=root.querySelector('.mm-rul__bar');
    var inchEl=root.querySelector('.mm-rul__inch');
    var cmEl=root.querySelector('.mm-rul__cm');
    if(!range)return;
    var min=num(root.getAttribute('data-min'),4.5),max=num(root.getAttribute('data-max'),7.0),def=num(root.getAttribute('data-def'),5.5);
    range.min=min;range.max=max;range.step=0.1;range.value=Math.min(max,Math.max(min,def));
    function upd(){
      var v=parseFloat(range.value);
      var p=max>min?(v-min)/(max-min):0;
      if(bar)bar.style.width=(20+p*80)+'%';
      if(inchEl)inchEl.textContent=v.toFixed(1)+'″';
      if(cmEl)cmEl.textContent='('+(v*2.54).toFixed(1)+'cm)';
    }
    range.addEventListener('input',upd);upd();
  }
  function init(){var l=document.querySelectorAll('.mm-rul');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
