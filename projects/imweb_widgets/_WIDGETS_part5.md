# 🧩 MAMORU 아임웹 위젯 — Part 5/5 (t08_typing ~ n14_feature_rows)

> 각 위젯=HTML/CSS/JS 3탭. 🚫 삼중괄호 {{{ }}} · CSS탭 {{변수}} · 인라인 style="…{{}}…"(동적값은 data-*+JS) · 인라인 on*= 금지. 이 파일 6종.

## 📑 이 파일의 위젯
- T8. 타이핑 헤드라인 — `t08_typing`
- T9. 읽기 진행바 + 맨위로 — `t09_progress`
- T10. 그립/손크기 시뮬레이터 — `t10_grip_simulator`
- T11. 롤링 로고 띠 (Marquee) — `t11_logo_marquee`
- N13. 이벤트 목록 카드 — `n13_event_cards`
- N14. 이미지+텍스트 교차 진열 — `n14_feature_rows`

---

## T8. 타이핑 헤드라인
`폴더: t08_typing`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "고정 문구+바뀌는 단어가 타이핑되며 순환." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 타이핑 헤드라인
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 고정 문구 + 바뀌는 단어가 타이핑되며 순환 (슬로건 임팩트)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name prefix @type outlined-textfield @default "마모루는" @label "앞 고정 문구" --}}
{{!-- @name suffix @type outlined-textfield @default "입니다" @label "뒤 고정 문구(선택)" --}}
{{!-- @name align @type outlined-textfield @default "가운데" @label "정렬 — 입력: 가운데 · 왼쪽" --}}
{{!-- @name phrases @type item @label "바뀌는 단어" --}}
<div class="mm-tw" data-align="{{align}}">
  <h2 class="mm-tw__line">
    <span class="mm-tw__prefix">{{prefix}}</span>
    <span class="mm-tw__word"></span><span class="mm-tw__cursor" aria-hidden="true"></span>
    <span class="mm-tw__suffix">{{suffix}}</span>
  </h2>
  <div class="mm-tw__data" hidden>
  {{#each phrases}}
    {{!-- @name phrase @type outlined-textfield @default "정직" @label "단어" --}}
    <span data-w="{{phrase}}"></span>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-tw{max-width:880px;margin:0 auto;padding:clamp(20px,4vw,32px) 16px;text-align:center;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-tw[data-align="왼쪽"]{text-align:left;}
.mm-tw__line{margin:0;font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:clamp(24px,7vw,46px);font-weight:900;letter-spacing:-.02em;line-height:1.25;color:#1A1A1A;}
.mm-tw__word{color:#1A1A1A;}
.mm-tw__cursor{display:inline-block;width:3px;height:1em;background:#1A1A1A;margin-left:2px;vertical-align:text-bottom;animation:mm-tw-blink 1s step-end infinite;}
@keyframes mm-tw-blink{0%,100%{opacity:1}50%{opacity:0}}
.mm-tw__suffix:empty{display:none;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-tw{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
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
```

---

## T9. 읽기 진행바 + 맨위로
`폴더: t09_progress`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "페이지 스크롤 진행바+맨위로(페이지당 1개만)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 읽기 진행바 + 맨위로
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭 (페이지 어디든 1개)
  📝 페이지 스크롤 진행도를 상단 바로 + 맨위로 버튼 (긴 페이지 UX)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name position @type outlined-textfield @default "상단" @label "바 위치 — 입력: 상단 · 하단" --}}
{{!-- @name thickness @type outlined-textfield @default "4" @label "바 두께(px)" --}}
{{!-- @name backTop @type switch @default true @label "맨위로 버튼" --}}
<div class="mm-prog" data-pos="{{position}}" data-thick="{{thickness}}" data-backtop="{{backTop}}">
  <div class="mm-prog__bar"></div>
  <button type="button" class="mm-prog__top" hidden aria-label="맨 위로 이동">↑</button>
</div>
```
### CSS 탭
```css
.mm-prog__bar{position:fixed;left:0;top:0;width:0;height:4px;background:#1A1A1A;z-index:99998;transition:width .1s linear;pointer-events:none;}
.mm-prog[data-pos="하단"] .mm-prog__bar{top:auto;bottom:0;}
.mm-prog__top{position:fixed;right:18px;bottom:18px;width:46px;height:46px;border:none;border-radius:50%;background:#1A1A1A;color:#FAF9F7;font-size:20px;line-height:1;cursor:pointer;z-index:99998;box-shadow:0 6px 20px rgba(0,0,0,.2);opacity:0;transform:translateY(10px);transition:opacity .3s,transform .3s;font-family:-apple-system,sans-serif;}
.mm-prog__top.is-show{opacity:1;transform:none;}
@media (hover:hover){.mm-prog__top:hover{background:#2D2D2D;}}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-prog{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var bar=root.querySelector('.mm-prog__bar');
    var topBtn=root.querySelector('.mm-prog__top');
    var thick=parseFloat(root.getAttribute('data-thick'));
    if(bar&&!isNaN(thick)&&thick>0)bar.style.height=thick+'px';
    var useTop=root.getAttribute('data-backtop')!=='false';
    if(!useTop&&topBtn)topBtn.style.display='none';
    function onScroll(){
      var de=document.documentElement,b=document.body;
      var st=window.pageYOffset||de.scrollTop||b.scrollTop||0;
      var h=(de.scrollHeight||b.scrollHeight)-window.innerHeight;
      var p=h>0?Math.min(st/h,1):0;
      if(bar)bar.style.width=(p*100)+'%';
      if(useTop&&topBtn)topBtn.classList.toggle('is-show',st>window.innerHeight*0.6);
    }
    if(useTop&&topBtn)topBtn.addEventListener('click',function(){window.scrollTo({top:0,behavior:'smooth'});});
    window.addEventListener('scroll',onScroll,{passive:true});
    window.addEventListener('resize',onScroll);onScroll();
  }
  function init(){var l=document.querySelectorAll('.mm-prog');if(!l.length){return setTimeout(init,50);}initOne(l[0]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## T10. 그립/손크기 시뮬레이터
`폴더: t10_grip_simulator`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "손 크기 선택 시 추천 가위 크기·모델." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 그립/손크기 시뮬레이터
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 손 크기 선택 → 추천 가위 크기(인치)·이유·모델 (전환)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name question @type outlined-textfield @default "손 크기를 선택하세요" @label "질문" --}}
{{!-- @name l1 @type outlined-textfield @default "작은 손" @label "옵션1 라벨" --}}
{{!-- @name l2 @type outlined-textfield @default "보통" @label "옵션2 라벨" --}}
{{!-- @name l3 @type outlined-textfield @default "큰 손" @label "옵션3 라벨" --}}
{{!-- @name ctaText @type outlined-textfield @default "추천 모델 보기" @label "버튼 문구" --}}
{{!-- @name results @type item @label "추천(옵션1·2·3 순서로 3개)" --}}
<div class="mm-grip" data-l1="{{l1}}" data-l2="{{l2}}" data-l3="{{l3}}" data-cta="{{ctaText}}">
  <p class="mm-grip__q">{{question}}</p>
  <div class="mm-grip__seg"></div>
  <div class="mm-grip__results">
  {{#each results}}
    {{!-- @name size @type outlined-textfield @default "5.5인치" @label "추천 크기" --}}
    {{!-- @name desc @type text-editor @default "<p>추천 이유</p>" @label "설명" --}}
    {{!-- @name link @type outlined-textfield @default "" @label "링크" --}}
    <div class="mm-grip__result" hidden>
      <span class="mm-grip__size">{{size}}</span>
      <div class="mm-grip__desc">{{desc}}</div>
      <a class="mm-grip__cta" href="{{link}}"></a>
    </div>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-grip{max-width:520px;margin:0 auto;padding:clamp(24px,5vw,36px) 20px;text-align:center;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-grip__q{margin:0 0 18px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(18px,4.6vw,24px);font-weight:800;letter-spacing:-.02em;}
.mm-grip__seg{display:inline-flex;padding:4px;background:#F5F3F0;border-radius:999px;gap:2px;}
.mm-grip__sbtn{appearance:none;cursor:pointer;border:none;background:none;padding:10px 20px;border-radius:999px;font-family:inherit;font-size:clamp(13px,3.6vw,15px);font-weight:700;color:#8A8580;transition:all .25s;}
.mm-grip__sbtn.is-on{background:#1A1A1A;color:#FAF9F7;}
.mm-grip__result{margin-top:26px;animation:mmgr .4s cubic-bezier(.4,0,.2,1) both;}
@keyframes mmgr{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
.mm-grip__size{display:block;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(40px,12vw,68px);font-weight:900;letter-spacing:-.03em;line-height:1;}
.mm-grip__desc{margin-top:12px;font-size:clamp(14px,3.8vw,16px);line-height:1.6;color:#4A4A4A;}
.mm-grip__desc p{margin:0 0 .15em;}
.mm-grip__cta{display:inline-flex;margin-top:18px;padding:13px 30px;border-radius:8px;background:#1A1A1A;color:#FAF9F7;font-weight:700;font-size:15px;text-decoration:none;transition:opacity .2s;}
.mm-grip__cta:empty{display:none;}
@media (hover:hover){.mm-grip__cta:hover{opacity:.88;}}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-grip{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var seg=root.querySelector('.mm-grip__seg');
    var results=root.querySelectorAll('.mm-grip__result');
    var ctaText=root.getAttribute('data-cta')||'';
    // CTA 채우기/숨김
    for(var r=0;r<results.length;r++){var a=results[r].querySelector('.mm-grip__cta');if(a){var h=a.getAttribute('href');if(!h)a.style.display='none';else a.textContent=ctaText;}}
    var labels=[root.getAttribute('data-l1'),root.getAttribute('data-l2'),root.getAttribute('data-l3')];
    var btns=[];
    for(var i=0;i<3;i++){if(!labels[i])continue;(function(idx){var b=document.createElement('button');b.type='button';b.className='mm-grip__sbtn';b.textContent=labels[idx];b.addEventListener('click',function(){select(idx);});seg.appendChild(b);btns.push({el:b,idx:idx});})(i);}
    function select(i){
      for(var k=0;k<btns.length;k++)btns[k].el.classList.toggle('is-on',btns[k].idx===i);
      for(var r=0;r<results.length;r++)results[r].hidden=(r!==i);
    }
    if(btns.length){select(btns[0].idx);}
  }
  function init(){var l=document.querySelectorAll('.mm-grip');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## T11. 롤링 로고 띠 (Marquee)
`폴더: t11_logo_marquee`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "로고 이미지가 가로로 끊김없이 흐르는 띠(마퀴). 아임웹 기본 '롤링 로고' 대체 + 가로 최대폭 지정 가능." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 롤링 로고 띠 (Marquee)
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 로고 이미지들이 가로로 끊김없이 흐르는 띠. 가로영역 확장해도 설정 폭에서 멈춤
  🚫 fetch·iframe·인라인핸들러 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name maxw @type outlined-textfield @default "" @label "가로 최대폭 — 숫자 자유 입력: 예 1280 (비우거나 full=꽉 채움). 영역 확장해도 이 값에서 멈춤" --}}
{{!-- @name bg @type color-picker @default "#5F5F5FFF" @label "배경 컬러" --}}
{{!-- @name size @type outlined-textfield @default "20" @label "로고 크기(px) — 로고 높이" --}}
{{!-- @name gap @type outlined-textfield @default "48" @label "로고 간 간격(px)" --}}
{{!-- @name padY @type outlined-textfield @default "10" @label "내부 상하 여백(px)" --}}
{{!-- @name speed @type outlined-textfield @default "30" @label "전환 속도(초) — 한 바퀴 도는 시간. 클수록 느림" --}}
{{!-- @name dir @type outlined-textfield @default "좌" @label "방향 — 입력: 좌 · 우" --}}
{{!-- @name gray @type switch @default false @label "흑백 이미지" --}}
{{!-- @name pause @type switch @default false @label "마우스 오버 시 정지" --}}
{{!-- @name logos @type item @label "롤링 로고" --}}
<div class="mm-lb" data-maxw="{{maxw}}" data-bg="{{bg}}" data-size="{{size}}" data-gap="{{gap}}" data-pady="{{padY}}" data-speed="{{speed}}" data-dir="{{dir}}" data-gray="{{gray}}" data-pause="{{pause}}">
  <div class="mm-lb__track">
    <div class="mm-lb__set">
      {{#each logos}}
        {{!-- @name logo @type image @label "로고 이미지 — 배경 투명 PNG 권장(높이 40px↑)" --}}
        <img class="mm-lb__logo" src="{{logo}}" alt="">
      {{/each}}
    </div>
  </div>
</div>
```
### CSS 탭
```css
.mm-lb{--lb-size:20px;--lb-gap:48px;--lb-pady:10px;--lb-speed:30s;--lb-bg:#5F5F5F;
  box-sizing:border-box;margin-left:auto;margin-right:auto;position:relative;overflow:hidden;background:var(--lb-bg);padding:var(--lb-pady) 0;}
.mm-lb__track{display:flex;width:max-content;will-change:transform;animation:mm-lb-scroll var(--lb-speed) linear infinite;}
/* half = 반복 단위(화면폭 이상). track에 half 2벌 → translateX(-50%)로 끊김없는 루프 */
.mm-lb__half{display:flex;align-items:center;flex:0 0 auto;}
/* 세트 하나 = 로고 묶음. trailing gap(padding-right)으로 세트·half 이음새 간격까지 동일 유지 */
.mm-lb__set{display:flex;align-items:center;gap:var(--lb-gap);padding-right:var(--lb-gap);flex:0 0 auto;}
.mm-lb__logo{height:var(--lb-size);width:auto;flex:0 0 auto;display:block;object-fit:contain;}
/* 방향: 우(→)면 역방향 */
.mm-lb[data-dir="우"] .mm-lb__track,.mm-lb[data-dir="right"] .mm-lb__track,.mm-lb[data-dir="→"] .mm-lb__track{animation-direction:reverse;}
/* 흑백 토글(switch → data-gray="true") */
.mm-lb[data-gray="true"] .mm-lb__logo{filter:grayscale(1);}
/* 마우스 오버 시 정지(switch → data-pause="true") */
.mm-lb[data-pause="true"]:hover .mm-lb__track{animation-play-state:paused;}
@keyframes mm-lb-scroll{from{transform:translateX(0)}to{transform:translateX(-50%)}}
@media (prefers-reduced-motion:reduce){.mm-lb__track{animation:none;}}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-lb{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
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
```

---

## N13. 이벤트 목록 카드
`폴더: n13_event_cards`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "이벤트 썸네일 카드 목록. 상태(진행중/종료) 뱃지, 클릭 시 링크 이동. PC 그리드 / 모바일 1열." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 이벤트 목록 카드
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭 (이벤트 목록 페이지에 삽입)
  📝 이벤트 썸네일+제목+상태+기간 카드. 클릭 → 이벤트 게시글/페이지로 이동
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "" @label "섹션 제목(선택) — 예: 진행중 이벤트" --}}
{{!-- @name theme @type outlined-textfield @default "라이트" @label "테마 — 입력: 라이트 · 다크(어두운 배경)" --}}
{{!-- @name cols @type outlined-textfield @default "3" @label "PC 한 줄 개수 — 예: 2 · 3 · 4" --}}
{{!-- @name maxw @type outlined-textfield @default "" @label "가로 최대폭 — 숫자 자유 입력(예 1280 · 비우면 꽉 채움). 영역 확장해도 이 값에서 멈춤" --}}
{{!-- @name events @type item @label "이벤트" --}}
<div class="mm-ev" data-theme="{{theme}}" data-cols="{{cols}}" data-maxw="{{maxw}}">
  <p class="mm-ev__title">{{title}}</p>
  <div class="mm-ev__grid">
  {{#each events}}
    {{!-- @name image @type image @label "썸네일 — 권장 1200×750px (16:10)" --}}
    {{!-- @name name @type outlined-textfield @default "이벤트 제목" @label "제목" --}}
    {{!-- @name status @type outlined-textfield @default "진행중" @label "상태 — 입력: 진행중 · 종료 · 예정" --}}
    {{!-- @name period @type outlined-textfield @default "" @label "기간(선택) — 예: 6.1 – 6.30" --}}
    {{!-- @name desc @type outlined-textfield @default "" @label "한 줄 설명(선택)" --}}
    {{!-- @name link @type outlined-textfield @default "" @label "링크 — 이벤트 게시글/페이지 URL" --}}
    <a class="mm-ev__card" href="{{link}}" data-status="{{status}}">
      <span class="mm-ev__thumb">
        <img class="mm-ev__img" src="{{image}}" alt="">
        <span class="mm-ev__badge">{{status}}</span>
      </span>
      <span class="mm-ev__body">
        <span class="mm-ev__name">{{name}}</span>
        <span class="mm-ev__period">{{period}}</span>
        <span class="mm-ev__desc">{{desc}}</span>
      </span>
    </a>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-ev{--ev-cols:3;box-sizing:border-box;max-width:1080px;margin:0 auto;padding:clamp(8px,2vw,16px) 0;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-ev[data-theme="다크"]{color:#FAF9F7;}
.mm-ev__title{margin:0 0 clamp(14px,3vw,22px);padding:0 4px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(18px,4vw,26px);font-weight:800;letter-spacing:-.02em;color:inherit;}
.mm-ev__title:empty{display:none;}
/* PC: 한 줄 N개(--ev-cols) 그리드. 모바일: 1열 */
.mm-ev__grid{display:grid;grid-template-columns:repeat(var(--ev-cols),minmax(0,1fr));gap:clamp(14px,2.5vw,22px);}
.mm-ev__card{display:flex;flex-direction:column;text-decoration:none;color:inherit;background:#FFFFFF;border:1px solid #EDEBE8;border-radius:14px;overflow:hidden;transition:transform .25s cubic-bezier(.4,0,.2,1),box-shadow .3s;}
.mm-ev[data-theme="다크"] .mm-ev__card{background:#2D2D2D;border-color:#3A3A3A;}
@media (hover:hover){.mm-ev__card:hover{transform:translateY(-3px);box-shadow:0 10px 30px rgba(0,0,0,.10);}.mm-ev__card:hover .mm-ev__img{transform:scale(1.05);}}
.mm-ev__thumb{position:relative;display:block;aspect-ratio:16/10;overflow:hidden;background:#F5F3F0;}
.mm-ev__img{width:100%;height:100%;object-fit:cover;display:block;transition:transform .5s cubic-bezier(.4,0,.2,1);}
/* 상태 뱃지(모노크롬): 진행중=Void채움 / 종료=Sand흐림 / 예정=아웃라인 */
.mm-ev__badge{position:absolute;top:10px;left:10px;padding:4px 11px;border-radius:999px;font-size:11px;font-weight:800;letter-spacing:.02em;background:#1A1A1A;color:#FAF9F7;}
.mm-ev__card[data-status="종료"] .mm-ev__badge{background:#D4D0CB;color:#4A4A4A;}
.mm-ev__card[data-status="예정"] .mm-ev__badge{background:#FAF9F7;color:#1A1A1A;box-shadow:inset 0 0 0 1px #1A1A1A;}
/* 종료 이벤트=썸네일 흑백으로 지난 느낌 */
.mm-ev__card[data-status="종료"] .mm-ev__img{filter:grayscale(1);opacity:.85;}
.mm-ev__body{display:flex;flex-direction:column;gap:4px;padding:clamp(12px,2vw,16px);}
.mm-ev__name{font-size:clamp(15px,2.4vw,17px);font-weight:800;letter-spacing:-.01em;line-height:1.35;color:inherit;}
.mm-ev__period{font-size:12px;font-weight:600;color:#8A8580;}
.mm-ev__period:empty{display:none;}
.mm-ev__desc{font-size:13px;line-height:1.5;color:#4A4A4A;}
.mm-ev[data-theme="다크"] .mm-ev__desc{color:#D4D0CB;}
.mm-ev__desc:empty{display:none;}
@media (max-width:768px){
  .mm-ev{padding-left:16px;padding-right:16px;}
  .mm-ev__grid{grid-template-columns:1fr;gap:14px;}
}
/* 등장(카드 순차 페이드) */
@media (prefers-reduced-motion:no-preference){
  .mm-ev__card{opacity:0;transform:translateY(16px);transition:opacity .5s cubic-bezier(.4,0,.2,1),transform .5s cubic-bezier(.4,0,.2,1),box-shadow .3s;}
  .mm-ev__card.is-in{opacity:1;transform:none;}
}
```
### JS 탭
```js
(function(){
  /* 가로 최대폭: 자유 숫자값(1280 등). 비움/full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  function applySettings(root){
    var c=String(root.getAttribute('data-cols')||'').trim();
    if(/^[1-9]\d*$/.test(c)) root.style.setProperty('--ev-cols',c);
    applyMaxw(root);
  }
  function initOne(root){
    applySettings(root);
    /* PC 열수·최대폭 바뀌면 즉시 반영(편집기 실시간) */
    if('MutationObserver' in window){ new MutationObserver(function(){applySettings(root);}).observe(root,{attributes:true,attributeFilter:['data-cols','data-maxw']}); }
    /* 카드 순차 등장 */
    var cards=root.querySelectorAll('.mm-ev__card');
    function reveal(){ for(var k=0;k<cards.length;k++){ cards[k].style.transitionDelay=((k%12)*0.06)+'s'; cards[k].classList.add('is-in'); } }
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){reveal();io.disconnect();}});},{threshold:.1});
      io.observe(root);
    }else{ reveal(); }
  }
  function init(){var l=document.querySelectorAll('.mm-ev');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## N14. 이미지+텍스트 교차 진열
`폴더: n14_feature_rows`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "이미지+텍스트 행을 지그재그로 교차 진열. PC 좌우 배치(홀수=사진오른쪽/짝수=사진왼쪽), 모바일 상하 스택." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 이미지+텍스트 교차 진열 (Feature Rows)
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 서비스/특징 소개 행. PC 좌우 지그재그, 모바일 상하 스택. 소제목+제목+설명+버튼
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name theme @type outlined-textfield @default "다크" @label "테마 — 입력: 다크(어두운 배경) · 라이트(밝은 배경)" --}}
{{!-- @name maxw @type outlined-textfield @default "" @label "가로 최대폭(PC) — 숫자 자유 입력(예 1100 · 비우면 꽉 채움). 영역 확장해도 이 값에서 멈춤" --}}
{{!-- @name items @type item @label "행 (이미지+텍스트)" --}}
<div class="mm-fr" data-theme="{{theme}}" data-maxw="{{maxw}}">
{{#each items}}
  {{!-- @name image @type image @label "이미지 — 권장 1000×800px" --}}
  {{!-- @name kicker @type outlined-textfield @default "" @label "소제목(영문, 선택) — 예: HYDRO REPAIR" --}}
  {{!-- @name title @type text-editor @default "<p>제목을 입력하세요</p>" @label "제목" --}}
  {{!-- @name desc @type text-editor @default "<p></p>" @label "설명(선택)" --}}
  {{!-- @name btnText @type outlined-textfield @default "" @label "버튼 문구(선택)" --}}
  {{!-- @name btnLink @type outlined-textfield @default "" @label "버튼 링크" --}}
  <div class="mm-fr__row">
    <div class="mm-fr__media"><img class="mm-fr__img" src="{{image}}" alt=""></div>
    <div class="mm-fr__body">
      <span class="mm-fr__kicker">{{kicker}}</span>
      <div class="mm-fr__title" role="heading" aria-level="2">{{title}}</div>
      <div class="mm-fr__desc">{{desc}}</div>
      <a class="mm-fr__btn" href="{{btnLink}}">{{btnText}}</a>
    </div>
  </div>
{{/each}}
</div>
```
### CSS 탭
```css
/* 폭 잡히기 전(풀폭) 깜빡임 방지: 기본 투명 → JS가 폭 적용 후 is-ready로 페이드인 */
.mm-fr{--fr-gap:clamp(28px,5vw,64px);box-sizing:border-box;max-width:1100px;margin:0 auto;display:flex;flex-direction:column;gap:clamp(44px,7vw,88px);padding:clamp(20px,3vw,32px) 0;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#FAF9F7;opacity:0;transition:opacity .35s ease;}
.mm-fr.is-ready{opacity:1;}
.mm-fr[data-theme="라이트"]{color:#1A1A1A;}
/* 행: 1행=사진 오른쪽(글씨 왼쪽), 2행=사진 왼쪽(글씨 오른쪽) → 지그재그(nth-child) */
.mm-fr__row{display:flex;align-items:center;gap:var(--fr-gap);flex-direction:row-reverse;}
.mm-fr__row:nth-child(even){flex-direction:row;}
.mm-fr__media{flex:1 1 0;min-width:0;}
.mm-fr__img{width:100%;height:auto;display:block;border-radius:18px;background:#2D2D2D;}
.mm-fr__img[src=""]{display:none;}
.mm-fr__body{flex:1 1 0;min-width:0;}
.mm-fr__kicker{display:block;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(11px,1.4vw,13px);font-weight:700;letter-spacing:.16em;text-transform:uppercase;color:#8A8580;margin-bottom:14px;}
.mm-fr__kicker:empty{display:none;}
.mm-fr__title{margin:0 0 16px;font-family:'Outfit','Plus Jakarta Sans','Pretendard','Noto Sans KR',sans-serif;font-size:clamp(22px,3.4vw,34px);font-weight:900;line-height:1.25;letter-spacing:-.02em;color:inherit;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-fr__title p{margin:0 0 .12em;}
.mm-fr__desc{font-size:clamp(14px,1.8vw,16px);line-height:1.65;color:#B8B4AF;margin-bottom:26px;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-fr[data-theme="라이트"] .mm-fr__desc{color:#4A4A4A;}
.mm-fr__desc:empty{display:none;}
.mm-fr__desc p{margin:0 0 .3em;}
.mm-fr__btn{display:inline-flex;align-items:center;padding:14px 28px;border-radius:10px;background:#FAF9F7;color:#1A1A1A;font-weight:700;font-size:15px;text-decoration:none;transition:transform .2s,opacity .2s;}
.mm-fr__btn::after{content:'→';margin-left:10px;}
.mm-fr[data-theme="라이트"] .mm-fr__btn{background:#1A1A1A;color:#FAF9F7;}
.mm-fr__btn:empty{display:none;}
.mm-fr__btn:active{transform:scale(.97);}
@media (hover:hover){.mm-fr__btn:hover{opacity:.9;}}
/* 모바일: 상하 스택(이미지 위·텍스트 아래) + 좌우 여백(안정감) */
@media (max-width:768px){
  .mm-fr{padding-left:16px;padding-right:16px;gap:clamp(36px,10vw,56px);}
  .mm-fr__row,.mm-fr__row:nth-child(even){flex-direction:column;gap:clamp(16px,4vw,24px);}
  .mm-fr__media,.mm-fr__body{flex:1 1 auto;width:100%;}
}
```
### JS 탭
```js
(function(){
  /* 가로 최대폭: 자유 숫자값(1100 등). 비움/full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  function initOne(root){
    applyMaxw(root);
    /* 최대폭 바뀌면 즉시 반영(편집기 실시간) */
    if('MutationObserver' in window){ new MutationObserver(function(){applyMaxw(root);}).observe(root,{attributes:true,attributeFilter:['data-maxw']}); }
    /* 폭 적용된 "뒤에" 페이드인 → 풀폭→축소 깜빡임 제거. 안전망 900ms */
    root.classList.add('is-ready');
    setTimeout(function(){ root.classList.add('is-ready'); }, 900);
  }
  function init(){var l=document.querySelectorAll('.mm-fr');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---
