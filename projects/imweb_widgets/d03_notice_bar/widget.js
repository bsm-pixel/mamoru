(function(){
  function initOne(root){
    if(root.getAttribute('data-flow')!=='true')return;
    var inEl=root.querySelector('.mm-notice__in');
    if(!inEl)return;
    // 끊김 없는 흐름: 내용 복제(50% 지점 동일)
    inEl.innerHTML=inEl.innerHTML+'<span style="display:inline-block;width:48px"></span>'+inEl.innerHTML;
  }
  function init(){var l=document.querySelectorAll('.mm-notice');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
