(function(){
  function initOne(root){
    var items=root.querySelectorAll('.mm-acc__item');
    for(var i=0;i<items.length;i++)(function(it){
      var head=it.querySelector('.mm-acc__head'),panel=it.querySelector('.mm-acc__panel');
      head.addEventListener('click',function(){
        var open=it.classList.contains('is-open');
        // 단일 오픈(아코디언): 다른 것 닫기
        for(var k=0;k<items.length;k++){items[k].classList.remove('is-open');items[k].querySelector('.mm-acc__panel').style.maxHeight='0px';}
        if(!open){it.classList.add('is-open');panel.style.maxHeight=panel.scrollHeight+'px';}
      });
    })(items[i]);
  }
  function init(){var l=document.querySelectorAll('.mm-acc');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
