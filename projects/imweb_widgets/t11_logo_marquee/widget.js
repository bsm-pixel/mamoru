(function(){
  function px(v){v=String(v==null?'':v).trim(); if(!v)return null; if(String(parseFloat(v))===v)v+='px'; return v;}
  function sec(v){v=String(v==null?'':v).trim(); if(!v)return null; if(String(parseFloat(v))===v)v+='s'; return v;}
  /* 가로 최대폭: 자유 숫자값(1280 등) 적용. 비움 → 꽉 채움, full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  /* 숫자·컬러 설정은 CSS 변수로 주입(색/크기/간격/여백/속도). 방향·흑백·정지는 CSS 속성선택자가 담당 */
  function applySettings(root){
    var s=px(root.getAttribute('data-size')); if(s)root.style.setProperty('--lb-size',s);
    var g=px(root.getAttribute('data-gap')); if(g!=null)root.style.setProperty('--lb-gap',g);
    var p=px(root.getAttribute('data-pady')); if(p!=null)root.style.setProperty('--lb-pady',p);
    var sp=sec(root.getAttribute('data-speed')); if(sp)root.style.setProperty('--lb-speed',sp);
    var bg=String(root.getAttribute('data-bg')||'').trim(); if(bg)root.style.setProperty('--lb-bg',bg);
    applyMaxw(root);
  }
  /* 이미지 로드 완료(또는 1.6s) 후 콜백 → 정확한 폭 측정 */
  function whenReady(root,cb){
    var imgs=root.querySelectorAll('.mm-lb__logo'),n=imgs.length,done=0,fired=false;
    function fire(){ if(!fired){ fired=true; cb(); } }
    if(!n){ return fire(); }
    function chk(){ if(done>=n) fire(); }
    for(var i=0;i<n;i++){ var im=imgs[i];
      if(im.complete&&im.naturalWidth){ done++; }
      else { im.addEventListener('load',function(){done++;chk();}); im.addEventListener('error',function(){done++;chk();}); }
    }
    chk(); setTimeout(fire,1600);
  }
  /* 끊김없는 마퀴: 반복 단위(half)를 '화면폭 이상'으로 타일링한 뒤 2벌 → translateX(-50%)면 빈틈 없이 순환 */
  function buildTrack(root){
    var track=root.querySelector('.mm-lb__track'); if(!track)return;
    /* 원본 로고 HTML 1회 저장(이후 재구성에 재사용) */
    if(root._lbLogos==null){ var s0=track.querySelector('.mm-lb__set'); root._lbLogos=s0?s0.innerHTML:''; }
    var logos=root._lbLogos||'';
    if(!logos.replace(/\s/g,'')) return; // 로고 없음
    /* 1) 단일 세트로 폭 측정 */
    track.innerHTML='<div class="mm-lb__set">'+logos+'</div>';
    var set=track.querySelector('.mm-lb__set');
    var setW=set.getBoundingClientRect().width, contW=root.getBoundingClientRect().width||setW;
    if(setW<2){ return setTimeout(function(){buildTrack(root);},150); } // 이미지 로딩 전 등 측정 실패 → 재시도
    /* 2) 반복 단위가 화면폭 이상이 되도록 세트 반복수 계산(+1 여유) */
    var copies=Math.max(1, Math.ceil(contW/setW)+1);
    var one=''; for(var c=0;c<copies;c++) one+='<div class="mm-lb__set">'+logos+'</div>';
    /* 3) half 2벌 → -50% 순환. 세트 trailing gap 덕에 이음새 간격도 동일 */
    track.innerHTML='<div class="mm-lb__half">'+one+'</div><div class="mm-lb__half">'+one+'</div>';
  }
  function initOne(root){
    applySettings(root);
    /* 패널값 바뀌면 즉시 재적용. 폭에 영향 주는 값(크기·간격·최대폭)은 트랙 재구성 */
    if('MutationObserver' in window){
      new MutationObserver(function(muts){
        applySettings(root);
        for(var i=0;i<muts.length;i++){ var a=muts[i].attributeName; if(a==='data-size'||a==='data-gap'||a==='data-maxw'){ buildTrack(root); break; } }
      }).observe(root,{attributes:true,attributeFilter:['data-maxw','data-bg','data-size','data-gap','data-pady','data-speed']});
    }
    whenReady(root,function(){ buildTrack(root); });
    /* 창 크기 변경 시 반복수 재계산(반응형) */
    if(window.addEventListener){ var t; window.addEventListener('resize',function(){ clearTimeout(t); t=setTimeout(function(){buildTrack(root);},200); }); }
  }
  function init(){var l=document.querySelectorAll('.mm-lb');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
