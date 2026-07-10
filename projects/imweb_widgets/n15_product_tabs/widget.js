(function(){
  /* 가로 최대폭: 자유 숫자값(1080 등). 비움/full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  function clean(v){ return (v==null?'':String(v)).trim(); }
  /* id로 요소 찾기. 아임웹이 PC/모바일 DOM을 복제 렌더하는 경우가 있어 getElementById 대신 전부 수집 */
  function byId(id){
    id=clean(id).replace(/"/g,'');
    if(!id) return [];
    try{ return [].slice.call(document.querySelectorAll('[id="'+id+'"]')); }catch(e){ return []; }
  }
  /* 외부 DOM은 display만 토글 — 내용은 절대 건드리지 않음 */
  function show(els,on){ for(var i=0;i<els.length;i++){ els[i].style.display = on ? '' : 'none'; } }

  /* 상품진열 그리드(#container_<WID>) 안의 상품 카드들 */
  function productCards(gridWid){
    var out=[], conts=byId('container_'+clean(gridWid));
    for(var c=0;c<conts.length;c++){
      var kids=conts[c].children;
      for(var i=0;i<kids.length;i++){
        var k=kids[i];
        if(k.classList && (k.classList.contains('shop-item')||k.classList.contains('_shop_item'))) out.push(k);
      }
    }
    return out;
  }
  /* 카드의 카테고리 라벨(.cate-label) 텍스트 */
  function cardCat(card){
    var el=card.querySelector('.cate-label');
    return el ? clean(el.textContent) : '';
  }
  /* 입력한 ID가 무엇인지 판별.
     상품진열(쇼핑) = #WID + #container_WID(그리드)  /  쇼핑기획전 = #WID 안에 .owl-carousel, container 없음 */
  function diagnoseId(gridWid){
    var id=clean(gridWid);
    if(!id) return 'ID 비어있음';
    if(byId('container_'+id).length) return 'ok';           /* 상품진열 맞음 */
    var outer=byId(id);
    if(!outer.length) return '해당 ID의 위젯을 페이지에서 찾지 못함';
    if(outer[0].querySelector && outer[0].querySelector('.owl-carousel')) return '이건 쇼핑기획전(캐러셀) ID입니다 — 쇼핑(상품진열) 위젯 ID를 넣으세요';
    return '이 ID엔 상품 그리드(#container_)가 없습니다 — 쇼핑(상품진열) 위젯인지 확인하세요';
  }

  function initOne(root){
    applyMaxw(root);
    /* 최대폭 바뀌면 즉시 반영(편집기 실시간) */
    if('MutationObserver' in window){ new MutationObserver(function(){applyMaxw(root);}).observe(root,{attributes:true,attributeFilter:['data-maxw']}); }

    var tabs=[].slice.call(root.querySelectorAll('.mm-pt__tab'));
    if(!tabs.length) return;
    var dbg=root.querySelector('.mm-pt__debug');
    var filterMode=(clean(root.getAttribute('data-mode'))==='상품필터');
    var gridWid=clean(root.getAttribute('data-gridwid'));

    function paintTabs(idx){
      for(var i=0;i<tabs.length;i++){
        var on=(i===idx);
        if(on) tabs[i].classList.add('is-on'); else tabs[i].classList.remove('is-on');
        tabs[i].setAttribute('aria-selected', on?'true':'false');
      }
    }
    /* [모드 A] 위젯전환 — 탭의 #WID 만 표시 */
    function activateSwitch(idx){
      paintTabs(idx);
      for(var i=0;i<tabs.length;i++) show(byId(tabs[i].getAttribute('data-wid')), i===idx);
    }
    /* [모드 B] 상품필터 — 상품진열 1개 안에서 카테고리 라벨로 카드 걸러내기 */
    function activateFilter(idx){
      paintTabs(idx);
      var want=clean(tabs[idx].getAttribute('data-cat'));   /* 비우면 '전체' */
      var cards=productCards(gridWid);
      for(var i=0;i<cards.length;i++){
        show([cards[i]], !want || cardCat(cards[i])===want);
      }
    }

    for(var i=0;i<tabs.length;i++){
      (function(k){ tabs[k].addEventListener('click', function(){ filterMode?activateFilter(k):activateSwitch(k); }); })(i);
    }

    /* 상품진열은 늦게 렌더될 수 있어 찾을 때까지 유한 폴링(최대 약 10초).
       body 감시자 미사용 — 아임웹 렌더와 경합 방지 */
    var tries=0;
    function resolve(){
      if(filterMode){
        var cards=productCards(gridWid);
        if(cards.length){
          /* 실제로 존재하는 카테고리 라벨을 모아 진단에 노출 → 탭에 그대로 복사해 넣으면 됨 */
          var seen={}, order=[], noLabel=0;
          for(var i=0;i<cards.length;i++){
            var c=cardCat(cards[i]);
            if(!c){ noLabel++; continue; }
            if(!(c in seen)){ seen[c]=0; order.push(c); }
            seen[c]++;
          }
          if(dbg){
            var parts=order.map(function(c){ return c+'('+seen[c]+')'; });
            dbg.textContent='진단: 상품 '+cards.length+'개 · 카테고리 라벨 '+(order.length?parts.join(' · '):'없음 — 위젯 옵션에서 카테고리 라벨 표시를 켜세요')
              + (noLabel?' · 라벨없는 상품 '+noLabel+'개':'');
          }
          if(order.length){ activateFilter(0); return; }   /* 라벨 있음 → 첫 탭 필터 적용 */
          /* 🛡️ 라벨이 하나도 없음 → 아무것도 숨기지 않음 */
          paintTabs(0); return;
        }
        /* 카드 0개 → 왜 그런지(기획전 ID 오입력 등) 구체적으로 알려줌. 마지막 시도에서만 확정 출력 */
        if(dbg && tries>=49) dbg.textContent='진단: '+diagnoseId(gridWid);
      } else {
        var found=0, missing=[];
        for(var j=0;j<tabs.length;j++){
          var w=clean(tabs[j].getAttribute('data-wid'));
          if(w && byId(w).length) found++; else missing.push(w||'(ID 비어있음)');
        }
        if(dbg){
          dbg.textContent = found===tabs.length
            ? '진단: 상품진열 '+found+'/'+tabs.length+' 모두 찾음 — 정상 동작'
            : '진단: 상품진열 '+found+'/'+tabs.length+' 찾음 · 못 찾음: '+missing.join(', ');
        }
        if(found===tabs.length){ activateSwitch(0); return; }
        if(tries>=49 && found>0){ activateSwitch(0); return; }   /* 일부만 찾아도 동작 */
      }
      if(++tries<50){ setTimeout(resolve,200); return; }
      /* 🛡️ 끝까지 못 찾음 → 아무것도 숨기지 않음(상품 전부 그대로 노출) */
      paintTabs(0);
    }
    resolve();
  }
  function init(){var l=document.querySelectorAll('.mm-pt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
