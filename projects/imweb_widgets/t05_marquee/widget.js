(function(){
  function initOne(root){
    var track=root.querySelector('.mm-mq__track');
    if(!track)return;
    // 끊김 없는 루프: 내용 복제(50% 지점에서 동일)
    track.innerHTML=track.innerHTML+track.innerHTML;
    var sp=parseFloat(root.getAttribute('data-speed'));
    track.style.setProperty('--mq-dur',((isNaN(sp)||sp<=0)?30:sp)+'s');
  }
  function init(){var l=document.querySelectorAll('.mm-mq');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
