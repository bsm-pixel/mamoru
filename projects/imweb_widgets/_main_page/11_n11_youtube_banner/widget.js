(function(){
  function vid(u){
    u=String(u||'').trim();
    if(/^[\w-]{11}$/.test(u))return u; // 그냥 ID
    var m=u.match(/(?:youtu\.be\/|v=|\/embed\/|\/shorts\/)([\w-]{11})/);
    return m?m[1]:'';
  }
  function initOne(root){
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
