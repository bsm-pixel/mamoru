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
  function cardCat(card){
    var el=card.querySelector('.cate-label');
    return el ? clean(el.textContent) : '';
  }
  /* 입력한 ID가 무엇인지 판별.
     상품진열(쇼핑) = #WID + #container_WID(그리드)  /  쇼핑기획전 = #WID 안에 .owl-carousel, container 없음 */
  function diagnoseId(gridWid){
    var id=clean(gridWid);
    if(!id) return '상품진열 위젯 ID가 비어 있습니다';
    if(byId('container_'+id).length) return 'ok';
    var outer=byId(id);
    if(!outer.length) return '"'+id+'" 위젯을 이 페이지에서 찾지 못했습니다 (ID 오타이거나, 같은 페이지에 없음)';
    if(outer[0].querySelector && outer[0].querySelector('.owl-carousel')) return '"'+id+'"는 쇼핑기획전(캐러셀)입니다 — 쇼핑(상품진열) 위젯 ID를 넣으세요';
    return '"'+id+'"에 상품 그리드(#container_)가 없습니다 — 쇼핑(상품진열) 위젯인지 확인하세요';
  }
  /* 페이지에 실제로 존재하는 상품진열 후보 ID들 — 오타/혼동 시 바로 알려주기 위함 */
  function foundGrids(){
    var out=[];
    try{
      var all=document.querySelectorAll('[id^="container_"]');
      for(var i=0;i<all.length;i++){ out.push(all[i].id.replace(/^container_/,'')); }
    }catch(e){}
    return out;
  }

  function initOne(root){
    applyMaxw(root);
    if('MutationObserver' in window){ new MutationObserver(function(){applyMaxw(root);}).observe(root,{attributes:true,attributeFilter:['data-maxw']}); }

    var tabs=[].slice.call(root.querySelectorAll('.mm-pt__tab'));
    if(!tabs.length) return;
    var dbg=root.querySelector('.mm-pt__debug');
    var mode=clean(root.getAttribute('data-mode'));
    var gridWid=clean(root.getAttribute('data-gridwid'));
    /* 🔑 모드 판정: [상품진열 위젯 ID]를 채웠으면 = 상품필터 의도. 모드 기본값(위젯전환)에 발목잡히지 않게 함.
       ID를 안 넣고 모드만 '상품필터'로 둔 경우엔 페이지의 상품진열을 자동 탐색(딱 1개일 때). */
    var filterMode = !!gridWid || (mode==='상품필터');
    var autoWid='';
    if(filterMode && !gridWid){
      var g0=foundGrids();
      if(g0.length===1){ gridWid=g0[0]; autoWid=' (자동탐색)'; }
    }

    function say(msg){ if(dbg) dbg.textContent='진단: '+msg; }
    /* 켜자마자 위젯이 실제로 받은 값을 보여줌 (JS가 도는지·값이 들어왔는지 즉시 확인) */
    say('모드='+(filterMode?'상품필터':'위젯전환')+' · 상품진열ID='+(gridWid||'없음')+autoWid+' · 탭 '+tabs.length+'개 · 확인 중…');

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
      for(var i=0;i<cards.length;i++) show([cards[i]], !want || cardCat(cards[i])===want);
    }

    for(var i=0;i<tabs.length;i++){
      (function(k){ tabs[k].addEventListener('click', function(){ filterMode?activateFilter(k):activateSwitch(k); }); })(i);
    }

    /* 상품진열은 늦게 렌더될 수 있어 유한 폴링(최대 약 10초). body 감시자 미사용(아임웹 렌더와 경합 방지) */
    var tries=0, LAST=49;
    function resolve(){
      if(filterMode){
        /* 상품진열이 늦게 렌더되면 init 시점 자동탐색이 빈손일 수 있어 폴링 중에도 재시도 */
        if(!gridWid){
          var gN=foundGrids();
          if(gN.length===1){ gridWid=gN[0]; autoWid=' (자동탐색)'; }
        }
        var cards=productCards(gridWid);
        if(cards.length){
          var seen={}, order=[], noLabel=0;
          for(var i=0;i<cards.length;i++){
            var c=cardCat(cards[i]);
            if(!c){ noLabel++; continue; }
            if(!(c in seen)){ seen[c]=0; order.push(c); }
            seen[c]++;
          }
          if(order.length){
            say('상품 '+cards.length+'개 · 카테고리 라벨 '+order.map(function(c){return c+'('+seen[c]+')';}).join(' · ')
                + (noLabel?' · 라벨없는 상품 '+noLabel+'개':'') + ' → 이 이름을 탭에 그대로 넣으세요');
            activateFilter(0); return;
          }
          /* 🛡️ 라벨이 하나도 없음 → 아무것도 숨기지 않음 */
          say('상품 '+cards.length+'개를 찾았지만 카테고리 라벨(.cate-label)이 없습니다 — 상품진열 위젯 옵션에서 카테고리 라벨 표시를 켜세요');
          paintTabs(0); return;
        }
        if(tries>=LAST){
          var g=foundGrids();
          say(diagnoseId(gridWid) + (g.length?' · 이 페이지의 상품진열 ID: '+g.join(', '):' · 이 페이지엔 상품진열이 하나도 없습니다'));
          paintTabs(0); return;
        }
      } else {
        var found=0, missing=[];
        for(var j=0;j<tabs.length;j++){
          var w=clean(tabs[j].getAttribute('data-wid'));
          if(w && byId(w).length) found++; else missing.push(w||'(ID 비어있음)');
        }
        if(found===tabs.length){ say('상품진열 '+found+'/'+tabs.length+' 모두 찾음 — 정상 동작'); activateSwitch(0); return; }
        if(tries>=LAST){
          say('상품진열 '+found+'/'+tabs.length+' 찾음 · 못 찾음: '+missing.join(', ')
              + (gridWid?'':' · 상품진열 위젯이 1개뿐이면 [상품진열 위젯 ID]를 채우고 동작 방식을 상품필터로 바꾸세요'));
          if(found>0){ activateSwitch(0); return; }
          paintTabs(0); return;
        }
      }
      tries++; setTimeout(resolve,200);
    }
    resolve();
  }
  function init(){var l=document.querySelectorAll('.mm-pt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
