(function(){
  function isLeft(al){al=String(al||'').trim();var low=al.toLowerCase();return al==='true'||al.indexOf('왼')>=0||al.indexOf('좌')>=0||low.indexOf('left')>=0||low.indexOf('start')>=0;}
  function initOne(root){
    root.setAttribute('data-align', isLeft(root.getAttribute('data-align'))?'왼쪽':'가운데');
    var wordEl=root.querySelector('.mm-tw__word');
    var data=root.querySelectorAll('.mm-tw__data span');
    var words=[];for(var i=0;i<data.length;i++){var w=data[i].getAttribute('data-w');if(w)words.push(w);}
    if(!wordEl||!words.length)return;
    var wi=0,ci=0,deleting=false;
    function tick(){
      var w=words[wi];
      if(!deleting){
        ci++;wordEl.textContent=w.slice(0,ci);
        if(ci>=w.length){deleting=true;return setTimeout(tick,1300);}
        return setTimeout(tick,90);
      }else{
        ci--;wordEl.textContent=w.slice(0,ci);
        if(ci<=0){deleting=false;wi=(wi+1)%words.length;return setTimeout(tick,250);}
        return setTimeout(tick,45);
      }
    }
    tick();
  }
  function init(){var l=document.querySelectorAll('.mm-tw');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
