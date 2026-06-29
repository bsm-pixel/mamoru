(function(){
  function initOne(root){
    var seg=root.querySelector('.mm-grip__seg');
    var results=root.querySelectorAll('.mm-grip__result');
    var ctaText=root.getAttribute('data-cta')||'';
    // CTA 채우기/숨김
    for(var r=0;r<results.length;r++){var a=results[r].querySelector('.mm-grip__cta');if(a){var h=a.getAttribute('href');if(!h)a.style.display='none';else a.textContent=ctaText;}}
    var labels=[root.getAttribute('data-l1'),root.getAttribute('data-l2'),root.getAttribute('data-l3')];
    var btns=[];
    for(var i=0;i<3;i++){if(!labels[i])continue;(function(idx){var b=document.createElement('button');b.type='button';b.className='mm-grip__sbtn';b.textContent=labels[idx];b.addEventListener('click',function(){select(idx);});seg.appendChild(b);btns.push({el:b,idx:idx});})(i);}
    function select(i){
      for(var k=0;k<btns.length;k++)btns[k].el.classList.toggle('is-on',btns[k].idx===i);
      for(var r=0;r<results.length;r++)results[r].hidden=(r!==i);
    }
    if(btns.length){select(btns[0].idx);}
  }
  function init(){var l=document.querySelectorAll('.mm-grip');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
