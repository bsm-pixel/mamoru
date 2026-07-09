(function(){
  /* 가로 최대폭: 자유 숫자값(1100 등)을 콘텐츠(target)에 적용. 비움/full → 꽉 채움. 값은 root의 data-maxw에서 읽음 */
  function applyMaxw(root, target){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ target.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ target.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; target.style.maxWidth=mw; }
  }
  /* 리치에디터가 넣은 앞뒤 빈 문단(<p></p>·<p><br></p>) 제거 → 제목↔설명 유령 간격 방지. 문단 사이 의도한 빈 줄은 유지 */
  function trimEmpty(el){
    if(!el) return; var k;
    while((k=el.firstElementChild) && k.tagName==='P' && !k.textContent.trim() && !k.querySelector('img')) el.removeChild(k);
    while((k=el.lastElementChild) && k.tagName==='P' && !k.textContent.trim() && !k.querySelector('img')) el.removeChild(k);
  }
  function initOne(root){
    var inner=root.querySelector('.mm-fr__inner')||root;
    var tx=root.querySelectorAll('.mm-fr__title,.mm-fr__desc');
    for(var t=0;t<tx.length;t++) trimEmpty(tx[t]);
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
