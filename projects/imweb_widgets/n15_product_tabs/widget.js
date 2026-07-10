(function(){
  /* 가로 최대폭: 자유 숫자값(1080 등). 비움/full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  /* WID로 네이티브 상품진열 위젯 찾기.
     아임웹이 PC/모바일 DOM 트리를 복제 렌더하는 경우가 있어 getElementById 대신 전부 수집 */
  function targets(wid){
    wid=(wid==null?'':wid).trim();
    if(!wid) return [];
    try{ return [].slice.call(document.querySelectorAll('[id="'+wid.replace(/"/g,'')+'"]')); }catch(e){ return []; }
  }
  /* 외부 DOM은 display만 토글 — 내용은 절대 건드리지 않음 */
  function show(els,on){ for(var i=0;i<els.length;i++){ els[i].style.display = on ? '' : 'none'; } }

  function initOne(root){
    applyMaxw(root);
    /* 최대폭 바뀌면 즉시 반영(편집기 실시간) */
    if('MutationObserver' in window){ new MutationObserver(function(){applyMaxw(root);}).observe(root,{attributes:true,attributeFilter:['data-maxw']}); }

    var tabs=[].slice.call(root.querySelectorAll('.mm-pt__tab'));
    if(!tabs.length) return;
    var dbg=root.querySelector('.mm-pt__debug');

    function activate(idx){
      for(var i=0;i<tabs.length;i++){
        var on=(i===idx);
        if(on) tabs[i].classList.add('is-on'); else tabs[i].classList.remove('is-on');
        tabs[i].setAttribute('aria-selected', on?'true':'false');
        show(targets(tabs[i].getAttribute('data-wid')), on);
      }
    }
    for(var i=0;i<tabs.length;i++){
      (function(k){ tabs[k].addEventListener('click', function(){ activate(k); }); })(i);
    }

    /* 상품진열 위젯은 늦게 렌더될 수 있어 찾을 때까지 유한 폴링(최대 약 10초).
       감시자(MutationObserver on body) 미사용 — 아임웹 렌더와 경합 방지 */
    var tries=0;
    function resolve(){
      var found=0, missing=[];
      for(var i=0;i<tabs.length;i++){
        var w=(tabs[i].getAttribute('data-wid')||'').trim();
        if(w && targets(w).length) found++; else missing.push(w||'(ID 비어있음)');
      }
      if(dbg){
        dbg.textContent = found===tabs.length
          ? '진단: 상품진열 ' + found + '/' + tabs.length + ' 모두 찾음 — 정상 동작'
          : '진단: 상품진열 ' + found + '/' + tabs.length + ' 찾음 · 못 찾음: ' + missing.join(', ');
      }
      if(found===tabs.length){ activate(0); return; }   /* 전부 찾음 → 첫 탭 활성 */
      if(++tries<50){ setTimeout(resolve,200); return; }
      if(found>0){ activate(0); return; }               /* 일부만 찾아도 동작 */
      /* 🛡️ 하나도 못 찾음 → 아무것도 숨기지 않음(상품진열 전부 그대로 노출) */
    }
    resolve();
  }
  function init(){var l=document.querySelectorAll('.mm-pt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
