(function(){
  function initOne(root){
    var opts=root.querySelectorAll('.mm-dx__opt');
    var results=root.querySelectorAll('.mm-dx__result');
    var ctaText=root.getAttribute('data-cta')||'';
    // CTA 텍스트 채우기 + 빈 링크 숨김
    for(var r=0;r<results.length;r++){
      var a=results[r].querySelector('.mm-dx__cta');
      if(a){var href=a.getAttribute('href');if(!href){a.style.display='none';}else{a.textContent=ctaText;}}
    }
    function select(i){
      for(var k=0;k<results.length;k++)results[k].hidden=(k!==i);
      for(var j=0;j<opts.length;j++)opts[j].classList.toggle('is-on',j===i);
      if(results[i]&&results[i].scrollIntoView){try{results[i].scrollIntoView({behavior:'smooth',block:'nearest'});}catch(_){}}
    }
    for(var x=0;x<opts.length;x++)(function(idx){opts[idx].addEventListener('click',function(){select(idx);});})(x);
  }
  function init(){var l=document.querySelectorAll('.mm-dx');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
