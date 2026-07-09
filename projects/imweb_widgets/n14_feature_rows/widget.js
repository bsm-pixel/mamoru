(function(){
  /* 가로 최대폭: 자유 숫자값(1100 등)을 콘텐츠(target)에 적용. 비움/full → 꽉 채움. 값은 root의 data-maxw에서 읽음 */
  function applyMaxw(root, target){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ target.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ target.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; target.style.maxWidth=mw; }
  }
  /* 에디터가 설명/제목 앞뒤에 자동 삽입하는 빈 문단(<p></p>·<p><br></p>·공백만) 제거 → 유령 간격 제거.
     ⚠ 로드 시 1회만 실행(감시자 없음) → 편집 중 줄바꿈을 절대 건드리지 않음. 실제 글자 있는 문단은 손 안 댐 */
  function trimEmpty(root){
    var ps=root.querySelectorAll('.mm-fr__title p, .mm-fr__desc p'), i, p;
    for(i=ps.length-1;i>=0;i--){ p=ps[i];
      if(!(p.textContent||'').trim() && !p.querySelector('img') && p.parentNode) p.parentNode.removeChild(p); }
  }
  function initOne(root){
    var inner=root.querySelector('.mm-fr__inner')||root;
    trimEmpty(root);
    setTimeout(function(){ trimEmpty(root); }, 400);  /* 지연 렌더 1회 보정(감시자 아님) */
    applyMaxw(root, inner);
    /* 최대폭 바뀌면 즉시 반영(편집기 실시간) */
    if('MutationObserver' in window){ new MutationObserver(function(){applyMaxw(root, inner);}).observe(root,{attributes:true,attributeFilter:['data-maxw']}); }
    /* 폭 적용된 "뒤에" 페이드인 → 풀폭→축소 깜빡임 제거. 안전망 900ms */
    root.classList.add('is-ready');
    setTimeout(function(){ root.classList.add('is-ready'); }, 900);
  }
  function init(){var l=document.querySelectorAll('.mm-fr');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
