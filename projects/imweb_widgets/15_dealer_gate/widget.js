(function(){
  function initOne(root){
    var code=(root.getAttribute('data-code')||'').trim().toLowerCase();
    var lock=root.querySelector('.mm-gate__lock');
    var content=root.querySelector('.mm-gate__content');
    var input=root.querySelector('.mm-gate__input');
    var btn=root.querySelector('.mm-gate__btn');
    var err=root.querySelector('.mm-gate__err');
    if(!input||!btn)return;
    if(err)err.textContent=err.getAttribute('data-msg')||'';
    function check(){
      if(input.value.trim().toLowerCase()===code&&code){
        if(lock)lock.hidden=true;if(content)content.hidden=false;
      }else{ if(err)err.hidden=false; }
    }
    btn.addEventListener('click',check);
    input.addEventListener('keydown',function(e){if(e.key==='Enter')check();});
    input.addEventListener('input',function(){if(err)err.hidden=true;});
  }
  function init(){var l=document.querySelectorAll('.mm-gate');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
