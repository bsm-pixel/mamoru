(function(){
  function vid(u){
    u=String(u||'').trim();
    if(/^[\w-]{11}$/.test(u))return u; // 그냥 ID
    var m=u.match(/(?:youtu\.be\/|v=|\/embed\/|\/shorts\/)([\w-]{11})/);
    return m?m[1]:'';
  }
  /* 가로 최대폭: 자유 숫자값(1280 등). 비움/full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  function applySettings(root){
    var c=String(root.getAttribute('data-cols')||'').trim();
    if(/^[1-9]\d*$/.test(c)) root.style.setProperty('--yt-cols',c);
    applyMaxw(root);
  }
  function initOne(root){
    applySettings(root);
    /* PC 열수·최대폭 바뀌면 즉시 반영(편집기 실시간) */
    if('MutationObserver' in window){ new MutationObserver(function(){applySettings(root);}).observe(root,{attributes:true,attributeFilter:['data-cols','data-maxw']}); }
    var cards=root.querySelectorAll('.mm-yt__card');
    for(var i=0;i<cards.length;i++){
      var id=vid(cards[i].getAttribute('data-url'));
      var custom=cards[i].getAttribute('data-thumb');
      var img=cards[i].querySelector('.mm-yt__img');
      if(id){
        cards[i].setAttribute('href','https://www.youtube.com/watch?v='+id);
        if(img)img.src= (custom&&custom.length>4) ? custom : 'https://img.youtube.com/vi/'+id+'/hqdefault.jpg';
      } else if(custom&&custom.length>4&&img){ img.src=custom; }
    }
  }
  function init(){var l=document.querySelectorAll('.mm-yt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
