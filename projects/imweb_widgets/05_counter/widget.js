(function(){
  function fmt(n){return n.toLocaleString('en-US');}
  function run(root){
    var nums=root.querySelectorAll('.mm-stat__num');
    for(var i=0;i<nums.length;i++)(function(el){
      var target=parseFloat(String(el.getAttribute('data-target')).replace(/[^0-9.]/g,''))||0;
      var pre=el.getAttribute('data-prefix')||'',suf=el.getAttribute('data-suffix')||'';
      var dur=1400,t0=null;
      function step(ts){if(!t0)t0=ts;var p=Math.min((ts-t0)/dur,1);var eased=1-Math.pow(1-p,3);
        el.textContent=pre+fmt(Math.round(target*eased))+suf;
        if(p<1)requestAnimationFrame(step);}
      requestAnimationFrame(step);
    })(nums[i]);
  }
  function initOne(root){
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){root.classList.add('is-in');run(root);io.disconnect();}});},{threshold:.3});
      io.observe(root);
    }else{root.classList.add('is-in');run(root);}
  }
  function init(){var l=document.querySelectorAll('.mm-stat');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
