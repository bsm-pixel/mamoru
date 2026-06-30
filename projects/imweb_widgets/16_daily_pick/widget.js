(function(){
  function dayOfYear(d){var s=new Date(d.getFullYear(),0,0);return Math.floor((d-s)/86400000);}
  function initOne(root){
    var cards=root.querySelectorAll('.mm-daily__card');
    if(!cards.length)return;
    var ctaText=root.getAttribute('data-cta')||'';
    var idx=dayOfYear(new Date())%cards.length;
    for(var i=0;i<cards.length;i++)cards[i].hidden=(i!==idx);
    var a=cards[idx].querySelector('.mm-daily__cta');
    if(a){var h=a.getAttribute('href');if(!h)a.style.display='none';else a.textContent=ctaText;}
  }
  function init(){var l=document.querySelectorAll('.mm-daily');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
