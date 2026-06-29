/* MAMORU 커스텀 위젯 — [위젯명] / JavaScript 탭
   바닐라 IIFE · 전 인스턴스 init · data-* 읽기 · 금지 API(fetch/iframe/lib/storage) 0 */
(function(){
  function initOne(root){
    // var cfg = root.getAttribute('data-config');
    // 진입 모션
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('is-in');io.unobserve(e.target);}});},{threshold:.15});
      io.observe(root);
    } else { root.classList.add('is-in'); }
  }
  function init(){
    var list=document.querySelectorAll('.mm-w');
    if(!list.length){return setTimeout(init,50);}
    for(var i=0;i<list.length;i++) initOne(list[i]);
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',init); else init();
})();
