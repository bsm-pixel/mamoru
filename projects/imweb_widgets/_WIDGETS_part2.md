# 🧩 MAMORU 아임웹 위젯 — Part 2/5 (11_recommender ~ d04_quote_banner)

> 각 위젯=HTML/CSS/JS 3탭. 🚫 삼중괄호 {{{ }}} 금지·인라인 on*= 금지. 이 파일 10종.

## 📑 이 파일의 위젯
- 11. 가위 추천 진단 (가이드형) — `11_recommender`
- 12. 360° 회전 뷰어 — `12_360_viewer`
- 13. 마감/컬러 옵션 미리보기 — `13_option_preview`
- 14. 스크롤 절단 히어로 — `14_cut_hero`
- 15. 딜러/아카데미 게이트 — `15_dealer_gate`
- 16. 오늘의 추천 모델 — `16_daily_pick`
- D1. 시네마틱 이미지 배너 (Ken Burns) — `d01_cinematic_banner`
- D2. 스플릿 프로모 배너 — `d02_split_promo`
- D3. 슬림 공지 띠 — `d03_notice_bar`
- D4. 대형 브랜드 인용 배너 — `d04_quote_banner`

---

## 11. 가위 추천 진단 (가이드형)
`폴더: 11_recommender`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "선택지 클릭 시 맞춤 모델 추천 카드 노출(링크 연결)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 가위 추천 진단(가이드형)
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 질문 → 선택지 클릭 → 맞춤 모델 추천 카드 + CTA (전환 도구)
  🚫 fetch·iframe 0 (폼 제출 불가 → 추천은 링크로 연결)
═══════════════════════════════════════════════════════════════ -->
{{!-- @name question @type outlined-textfield @default "어떤 가위를 찾으세요?" @label "질문" --}}
{{!-- @name ctaText @type outlined-textfield @default "자세히 보기" @label "추천 버튼 문구" --}}
{{!-- @name options @type item @label "선택지 & 추천" --}}
<div class="mm-dx" data-cta="{{ctaText}}">
  <p class="mm-dx__q">{{question}}</p>
  <div class="mm-dx__opts">
  {{#each options}}
    {{!-- @name optLabel @type outlined-textfield @default "선택지" @label "선택지 버튼" --}}
    {{!-- @name recName @type outlined-textfield @default "추천 모델" @label "추천 모델명" --}}
    {{!-- @name recDesc @type text-editor @default "<p>추천 이유</p>" @label "추천 설명" --}}
    {{!-- @name recImage @type image @label "추천 이미지(선택) (권장 1600×900px)" --}}
    {{!-- @name recLink @type outlined-textfield @default "" @label "추천 링크" --}}
    <button type="button" class="mm-dx__opt">{{optLabel}}</button>
  {{/each}}
  </div>
  <div class="mm-dx__results">
  {{#each options}}
    <div class="mm-dx__result" hidden>
      <img class="mm-dx__img" src="{{recImage}}" alt="">
      <h3 class="mm-dx__name">{{recName}}</h3>
      <div class="mm-dx__desc">{{recDesc}}</div>
      <a class="mm-dx__cta" href="{{recLink}}"></a>
    </div>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-dx{max-width:600px;margin:0 auto;padding:clamp(24px,5vw,36px) 20px;text-align:center;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-dx__q{margin:0 0 22px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(19px,5vw,26px);font-weight:800;letter-spacing:-.02em;line-height:1.3;}
.mm-dx__opts{display:flex;flex-wrap:wrap;gap:10px;justify-content:center;}
.mm-dx__opt{appearance:none;cursor:pointer;padding:12px 22px;border-radius:999px;border:1px solid #D4D0CB;background:#FFFFFF;color:#1A1A1A;font-family:inherit;font-size:clamp(14px,3.6vw,15px);font-weight:600;transition:all .2s cubic-bezier(.4,0,.2,1);}
.mm-dx__opt:active{transform:scale(.97);}
.mm-dx__opt.is-on{background:#1A1A1A;color:#FAF9F7;border-color:#1A1A1A;}
@media (hover:hover){.mm-dx__opt:hover{border-color:#1A1A1A;}}
.mm-dx__result{margin-top:24px;padding:clamp(22px,5vw,32px);background:#FAF9F7;border:1px solid #D4D0CB;border-radius:16px;animation:mmdx-in .45s cubic-bezier(.4,0,.2,1) both;}
@keyframes mmdx-in{from{opacity:0;transform:translateY(16px);}to{opacity:1;transform:none;}}
.mm-dx__img{width:100%;aspect-ratio:16/10;object-fit:cover;border-radius:10px;background:#F5F3F0;margin-bottom:16px;display:block;}
.mm-dx__name{margin:0 0 10px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(18px,4.5vw,22px);font-weight:800;letter-spacing:-.02em;}
.mm-dx__desc{font-size:clamp(14px,3.8vw,16px);line-height:1.6;color:#4A4A4A;}
.mm-dx__desc p{margin:0 0 .15em;}
.mm-dx__cta{display:inline-flex;align-items:center;justify-content:center;margin-top:18px;padding:13px 30px;border-radius:8px;background:#1A1A1A;color:#FAF9F7;font-weight:700;font-size:15px;text-decoration:none;transition:opacity .2s;}
.mm-dx__cta:empty{display:none;}
@media (hover:hover){.mm-dx__cta:hover{opacity:.88;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var opts=root.querySelectorAll('.mm-dx__opt');
    var results=root.querySelectorAll('.mm-dx__result');
    var ctaText=root.getAttribute('data-cta')||'';
    // CTA 텍스트 채우기 + 빈 링크 숨김
    for(var r=0;r<results.length;r++){
      var a=results[r].querySelector('.mm-dx__cta');
      if(a){var href=a.getAttribute('href');if(!href){a.style.display='none';}else{a.textContent=ctaText;}}
    }
    function select(i){
      for(var k=0;k<results.length;k++)results[k].hidden=(k!==i);
      for(var j=0;j<opts.length;j++)opts[j].classList.toggle('is-on',j===i);
      if(results[i]&&results[i].scrollIntoView){try{results[i].scrollIntoView({behavior:'smooth',block:'nearest'});}catch(_){}}
    }
    for(var x=0;x<opts.length;x++)(function(idx){opts[idx].addEventListener('click',function(){select(idx);});})(x);
  }
  function init(){var l=document.querySelectorAll('.mm-dx');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## 12. 360° 회전 뷰어
`폴더: 12_360_viewer`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "각도별 사진을 드래그로 360° 회전." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 360° 회전 뷰어
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 각도별 사진들을 드래그로 빙글 회전 (제품 디테일 신뢰)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name ratio @type outlined-textfield @default "4/3" @label "이미지 비율 — 입력: 4/3 · 1/1 · 16/9" --}}
{{!-- @name frames @type item @label "회전 프레임(각도별 사진)" --}}
<div class="mm-360">
  <div class="mm-360__stage" style="aspect-ratio:{{ratio}};">
  {{#each frames}}
    {{!-- @name frame @type image @label "프레임 이미지 (권장 800×800px)" --}}
    <img class="mm-360__f" src="{{frame}}" alt="" draggable="false">
  {{/each}}
    <span class="mm-360__hint" aria-hidden="true">↔ 드래그하여 회전</span>
  </div>
</div>
```
### CSS 탭
```css
.mm-360{max-width:560px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-360__stage{position:relative;width:100%;aspect-ratio:4/3;overflow:hidden;border-radius:14px;background:#F5F3F0;touch-action:pan-y;cursor:ew-resize;user-select:none;}
.mm-360__f{position:absolute;inset:0;width:100%;height:100%;object-fit:contain;display:none;pointer-events:none;}
.mm-360__f.is-on{display:block;}
.mm-360__hint{position:absolute;left:50%;bottom:12px;transform:translateX(-50%);z-index:2;font-size:11px;font-weight:600;letter-spacing:.04em;color:#8A8580;background:rgba(250,249,247,.85);padding:5px 12px;border-radius:999px;pointer-events:none;transition:opacity .3s;}
.mm-360.is-active .mm-360__hint{opacity:0;}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var stage=root.querySelector('.mm-360__stage');
    var frames=root.querySelectorAll('.mm-360__f');
    var F=frames.length;
    if(!stage||!F)return;
    var idx=0,dragging=false,startX=0,startIdx=0;
    function show(i){idx=((i%F)+F)%F;for(var k=0;k<F;k++)frames[k].classList.toggle('is-on',k===idx);}
    function move(x){var step=stage.getBoundingClientRect().width/F||1;var d=Math.round((x-startX)/step);show(startIdx-d);}
    stage.addEventListener('pointerdown',function(e){dragging=true;root.classList.add('is-active');startX=e.clientX;startIdx=idx;try{stage.setPointerCapture(e.pointerId);}catch(_){}});
    stage.addEventListener('pointermove',function(e){if(dragging)move(e.clientX);});
    window.addEventListener('pointerup',function(){dragging=false;});
    show(0);
  }
  function init(){var l=document.querySelectorAll('.mm-360');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## 13. 마감/컬러 옵션 미리보기
`폴더: 13_option_preview`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "옵션 클릭 시 제품 이미지 즉시 교체." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 마감/컬러 옵션 미리보기
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 옵션 클릭 → 제품 이미지 즉시 교체 (구매 확신)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "" @label "제목(선택)" --}}
{{!-- @name ratio @type outlined-textfield @default "4/3" @label "이미지 비율 — 입력: 4/3 · 1/1 · 16/9" --}}
{{!-- @name options @type item @label "옵션" --}}
<div class="mm-opt">
  <p class="mm-opt__title">{{title}}</p>
  <div class="mm-opt__stage" style="aspect-ratio:{{ratio}};"><img class="mm-opt__main" src="" alt=""></div>
  <div class="mm-opt__btns">
  {{#each options}}
    {{!-- @name optName @type outlined-textfield @default "옵션" @label "옵션명" --}}
    {{!-- @name optImage @type image @label "옵션 이미지 (권장 800×800px)" --}}
    <button type="button" class="mm-opt__btn" data-img="{{optImage}}">{{optName}}</button>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-opt{max-width:560px;margin:0 auto;text-align:center;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-opt__title{margin:0 0 16px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(18px,4.5vw,22px);font-weight:800;letter-spacing:-.02em;}
.mm-opt__title:empty{display:none;}
.mm-opt__stage{position:relative;width:100%;aspect-ratio:4/3;overflow:hidden;border-radius:14px;background:#F5F3F0;}
.mm-opt__main{width:100%;height:100%;object-fit:cover;display:block;transition:opacity .3s cubic-bezier(.4,0,.2,1);}
.mm-opt__btns{display:flex;flex-wrap:wrap;gap:8px;justify-content:center;margin-top:16px;}
.mm-opt__btn{appearance:none;cursor:pointer;padding:10px 18px;border-radius:999px;border:1px solid #D4D0CB;background:#FFFFFF;color:#1A1A1A;font-family:inherit;font-size:14px;font-weight:600;transition:all .2s;}
.mm-opt__btn:active{transform:scale(.97);}
.mm-opt__btn.is-on{background:#1A1A1A;color:#FAF9F7;border-color:#1A1A1A;}
@media (hover:hover){.mm-opt__btn:hover{border-color:#1A1A1A;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var main=root.querySelector('.mm-opt__main');
    var btns=root.querySelectorAll('.mm-opt__btn');
    if(!main||!btns.length)return;
    function swap(i){
      var img=btns[i].getAttribute('data-img')||'';
      main.style.opacity='0';
      setTimeout(function(){main.src=img;main.style.opacity='1';},150);
      for(var k=0;k<btns.length;k++)btns[k].classList.toggle('is-on',k===i);
    }
    for(var x=0;x<btns.length;x++)(function(idx){btns[idx].addEventListener('click',function(){swap(idx);});})(x);
    // 초기: 첫 옵션
    main.src=btns[0].getAttribute('data-img')||'';btns[0].classList.add('is-on');
  }
  function init(){var l=document.querySelectorAll('.mm-opt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## 14. 스크롤 절단 히어로
`폴더: 14_cut_hero`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "진입 시 hairline 컷 후 헤드라인 등장(브랜드 히어로)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 스크롤 절단 히어로
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 진입 시 hairline이 가로질러 '컷' → 헤드라인 등장 (브랜드 첫인상)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name bg @type image @label "PC 배경 이미지(선택) — 권장 1600×900px" --}}
{{!-- @name bgMobile @type image @label "모바일 배경 이미지(선택) — 권장 1080×1350px · 비우면 PC 이미지 사용" --}}
{{!-- @name theme @type outlined-textfield @default "다크" @label "테마 — 입력: 다크 · 라이트" --}}
{{!-- @name height @type outlined-textfield @default "" @label "높이 — 예: 440px 또는 60vh (비우면 자동)" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name kicker @type outlined-textfield @default "CUT THE FAKE, KEEP THE REAL" @label "키커(영문)" --}}
{{!-- @name headline @type text-editor @default "<p>좋은 미용가위, 그 기준을 정의하다</p>" @label "헤드라인" --}}
{{!-- @name sub @type text-editor @default "<p></p>" @label "서브 문구(선택)" --}}
{{!-- @name btnText @type outlined-textfield @default "" @label "버튼 문구(선택)" --}}
{{!-- @name btnLink @type outlined-textfield @default "" @label "버튼 링크" --}}
<div class="mm-cut" data-theme="{{theme}}" data-height="{{height}}">
  <img class="mm-cut__bg" src="{{bg}}" alt="">
  <img class="mm-cut__bgm" src="{{bgMobile}}" alt="">
  <div class="mm-cut__inner">
    <span class="mm-cut__kicker">{{kicker}}</span>
    <span class="mm-cut__blade" aria-hidden="true"></span>
    <div class="mm-cut__headline" role="heading" aria-level="2">{{headline}}</div>
    <div class="mm-cut__sub">{{sub}}</div>
    <a class="mm-cut__btn" href="{{btnLink}}">{{btnText}}</a>
  </div>
</div>
```
### CSS 탭
```css
.mm-cut{position:relative;overflow:hidden;border-radius:{{radius}};min-height:clamp(280px,50vw,440px);display:flex;align-items:center;justify-content:center;text-align:center;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;background:#1A1A1A;}
.mm-cut[data-theme="라이트"]{background:#FAF9F7;}
.mm-cut__bg,.mm-cut__bgm{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;opacity:.5;z-index:0;}
.mm-cut__bgm{display:none;}
@media (max-width:768px){.mm-cut[data-hasm="1"] .mm-cut__bg{display:none;}.mm-cut[data-hasm="1"] .mm-cut__bgm{display:block;}}
.mm-cut[data-theme="라이트"] .mm-cut__bg,.mm-cut[data-theme="라이트"] .mm-cut__bgm{opacity:.85;}
.mm-cut__inner{position:relative;z-index:1;padding:clamp(28px,6vw,56px);max-width:760px;}
.mm-cut__kicker{display:block;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(10px,2.6vw,12px);font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#B8B4AF;opacity:0;transition:opacity .6s ease .1s;}
.mm-cut[data-theme="라이트"] .mm-cut__kicker{color:#8A8580;}
.mm-cut__blade{display:block;width:0;height:1px;margin:18px auto;background:#FAF9F7;transition:width .7s cubic-bezier(.4,0,.2,1) .15s;}
.mm-cut[data-theme="라이트"] .mm-cut__blade{background:#1A1A1A;}
.mm-cut__headline{margin:0;font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:clamp(24px,6.5vw,44px);font-weight:900;line-height:1.2;letter-spacing:-.02em;color:#FAF9F7;opacity:0;transform:translateY(20px);transition:opacity .7s cubic-bezier(.4,0,.2,1) .35s,transform .7s cubic-bezier(.4,0,.2,1) .35s;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-cut[data-theme="라이트"] .mm-cut__headline{color:#1A1A1A;}
.mm-cut__sub{margin:16px 0 0;font-size:clamp(14px,3.8vw,17px);line-height:1.6;color:#D4D0CB;opacity:0;transition:opacity .7s ease .55s;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-cut[data-theme="라이트"] .mm-cut__sub{color:#4A4A4A;}
.mm-cut__sub:empty{display:none;}
.mm-cut__btn{display:inline-flex;margin-top:26px;padding:14px 34px;border-radius:8px;background:#FAF9F7;color:#1A1A1A;font-weight:700;font-size:15px;text-decoration:none;opacity:0;transition:opacity .7s ease .7s,transform .2s;}
.mm-cut[data-theme="라이트"] .mm-cut__btn{background:#1A1A1A;color:#FAF9F7;}
.mm-cut__btn:empty{display:none;}
.mm-cut__btn:active{transform:scale(.97);}
.mm-cut.is-in .mm-cut__kicker,.mm-cut.is-in .mm-cut__headline,.mm-cut.is-in .mm-cut__sub,.mm-cut.is-in .mm-cut__btn{opacity:1;transform:none;}
.mm-cut.is-in .mm-cut__blade{width:clamp(40px,12vw,90px);}

/* 여러 줄(text-editor) 입력 시 문단 줄간격 통일 */
.mm-cut__headline p,.mm-cut__sub p{margin:0 0 .15em;}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var h=root.getAttribute('data-height');
    if(h){ var hv=h.trim(); if(hv){ if(String(parseFloat(hv))===hv) hv+='px'; root.style.minHeight=hv; } }
    var m=root.querySelector('.mm-cut__bgm');
    if(m && String(m.getAttribute('src')||'').trim()) root.setAttribute('data-hasm','1');
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('is-in');io.unobserve(e.target);}});},{threshold:.3});
      io.observe(root);
    }else{root.classList.add('is-in');}
  }
  function init(){var l=document.querySelectorAll('.mm-cut');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## 15. 딜러/아카데미 게이트
`폴더: 15_dealer_gate`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "코드 입력 시 전용 내용 노출(가벼운 구분용·보안 아님)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 딜러/아카데미 게이트
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 코드 입력 → 전용 안내/가격 노출 (B2B 정보 분리)
  🚫 fetch·iframe 0 / ⚠️ 클라이언트 코드라 강한 보안용 아님
═══════════════════════════════════════════════════════════════ -->
{{!-- @name prompt @type outlined-textfield @default "전용 코드를 입력하세요" @label "안내 문구" --}}
{{!-- @name code @type outlined-textfield @default "MAMORU" @label "통과 코드" --}}
{{!-- @name btnText @type outlined-textfield @default "확인" @label "버튼 문구" --}}
{{!-- @name wrongText @type outlined-textfield @default "코드가 올바르지 않습니다" @label "오류 문구" --}}
{{!-- @name content @type text-editor @default "<p>딜러/아카데미 전용 안내입니다.</p>" @label "통과 후 내용" --}}
<div class="mm-gate" data-code="{{code}}">
  <div class="mm-gate__lock">
    <p class="mm-gate__prompt">{{prompt}}</p>
    <div class="mm-gate__row">
      <input type="text" class="mm-gate__input" aria-label="전용 코드" autocomplete="off">
      <button type="button" class="mm-gate__btn">{{btnText}}</button>
    </div>
    <p class="mm-gate__err" data-msg="{{wrongText}}" hidden></p>
  </div>
  <div class="mm-gate__content" hidden>{{content}}</div>
</div>
```
### CSS 탭
```css
.mm-gate{max-width:440px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-gate__lock{padding:clamp(24px,5vw,32px);border:1px solid #D4D0CB;border-radius:16px;background:#FAF9F7;text-align:center;}
.mm-gate__prompt{margin:0 0 16px;font-size:clamp(15px,4vw,17px);font-weight:700;}
.mm-gate__row{display:flex;gap:8px;}
.mm-gate__input{flex:1;padding:13px 14px;border:1px solid #D4D0CB;border-radius:8px;background:#FFFFFF;font-family:inherit;font-size:15px;color:#1A1A1A;-webkit-appearance:none;}
.mm-gate__input:focus{outline:none;border-color:#1A1A1A;}
.mm-gate__btn{padding:13px 22px;border:none;border-radius:8px;background:#1A1A1A;color:#FAF9F7;font-family:inherit;font-weight:700;font-size:15px;cursor:pointer;}
.mm-gate__err{margin:12px 0 0;font-size:13px;color:#8A8580;}
.mm-gate__content{animation:mmgate .4s cubic-bezier(.4,0,.2,1) both;font-size:clamp(14px,3.8vw,16px);line-height:1.7;color:#2D2D2D;}
.mm-gate__content p{margin:0 0 .15em;}
@keyframes mmgate{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}}
```
### JS 탭
```js
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
```

---

## 16. 오늘의 추천 모델
`폴더: 16_daily_pick`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "날짜 기준 매일 다른 추천 모델 1개 노출." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 오늘의 추천 모델
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 날짜 기준으로 매일 다른 추천 모델 1개 노출 (재방문 유도)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name kicker @type outlined-textfield @default "오늘의 추천" @label "키커" --}}
{{!-- @name ctaText @type outlined-textfield @default "자세히 보기" @label "버튼 문구" --}}
{{!-- @name models @type item @label "추천 풀(모델들)" --}}
<div class="mm-daily" data-cta="{{ctaText}}">
  <span class="mm-daily__kicker">{{kicker}}</span>
  <div class="mm-daily__cards">
  {{#each models}}
    {{!-- @name name @type outlined-textfield @default "모델명" @label "모델명" --}}
    {{!-- @name image @type image @label "이미지 (권장 1600×900px)" --}}
    {{!-- @name desc @type text-editor @default "<p>설명</p>" @label "설명" --}}
    {{!-- @name link @type outlined-textfield @default "" @label "링크" --}}
    <article class="mm-daily__card" hidden>
      <img class="mm-daily__img" src="{{image}}" alt="">
      <div class="mm-daily__body">
        <h3 class="mm-daily__name">{{name}}</h3>
        <div class="mm-daily__desc">{{desc}}</div>
        <a class="mm-daily__cta" href="{{link}}"></a>
      </div>
    </article>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-daily{max-width:600px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;text-align:center;}
.mm-daily__kicker{display:inline-block;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:11px;font-weight:700;letter-spacing:.16em;text-transform:uppercase;color:#8A8580;margin-bottom:14px;}
.mm-daily__card{background:#FFFFFF;border:1px solid #D4D0CB;border-radius:16px;overflow:hidden;text-align:left;animation:mmdaily .5s cubic-bezier(.4,0,.2,1) both;}
@keyframes mmdaily{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
.mm-daily__img{width:100%;aspect-ratio:16/10;object-fit:cover;background:#F5F3F0;display:block;}
.mm-daily__body{padding:clamp(20px,4vw,28px);}
.mm-daily__name{margin:0 0 10px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(18px,4.5vw,24px);font-weight:800;letter-spacing:-.02em;}
.mm-daily__desc{font-size:clamp(14px,3.8vw,16px);line-height:1.6;color:#4A4A4A;}
.mm-daily__desc p{margin:0 0 .15em;}
.mm-daily__cta{display:inline-flex;margin-top:16px;padding:12px 26px;border-radius:8px;background:#1A1A1A;color:#FAF9F7;font-weight:700;font-size:14px;text-decoration:none;transition:opacity .2s;}
.mm-daily__cta:empty{display:none;}
@media (hover:hover){.mm-daily__cta:hover{opacity:.88;}}
```
### JS 탭
```js
(function(){
  function dayOfYear(d){var s=new Date(d.getFullYear(),0,0);return Math.floor((d-s)/86400000);}
  function initOne(root){
    var cards=root.querySelectorAll('.mm-daily__card');
    if(!cards.length)return;
    var ctaText=root.getAttribute('data-cta')||'';
    var idx=dayOfYear(new Date())%cards.length;
    for(var i=0;i<cards.length;i++)cards[i].hidden=(i!==idx);
    var a=cards[idx].querySelector('.mm-daily__cta');
    if(a){var h=a.getAttribute('href');if(!h)a.style.display='none';else a.textContent=ctaText;}
  }
  function init(){var l=document.querySelectorAll('.mm-daily');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## D1. 시네마틱 이미지 배너 (Ken Burns)
`폴더: d01_cinematic_banner`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "배경 느린 줌+카피 오버레이 고급 배너. 높이·모서리·사진초점 설정 가능." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 시네마틱 이미지 배너 (Ken Burns)
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 풀폭 이미지 느린 줌 + 카피·CTA 오버레이 (범용 고급 배너)
  🚫 fetch·iframe·인라인핸들러 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name image @type image @label "PC 배경 이미지 — 권장 1600×900px (가로형)" --}}
{{!-- @name imageMobile @type image @label "모바일 배경 이미지 (선택) — 권장 1080×1350px (세로형) · 비우면 PC 이미지 사용" --}}
{{!-- @name height @type outlined-textfield @default "" @label "높이 — 예: 480px 또는 60vh (비우면 자동)" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name maxw @type outlined-textfield @default "" @label "최대 가로폭(PC) — 예: 1000px (비우면 꽉 채움) · 모바일은 자동 꽉 채움" --}}
{{!-- @name focus @type outlined-textfield @default "중앙" @label "사진 초점 — 입력: 중앙·좌·우·상·하" --}}
{{!-- @name overlay @type color-picker @default "#1A1A1A80" @label "어둡게 (검정의 투명도 슬라이더를 드래그 · 맨뒤 2자리=어둡기)" --}}
{{!-- @name align @type outlined-textfield @default "가운데" @label "정렬 (가운데/왼쪽 · 유형을 옵션버튼/스위치로 바꿔도 됨)" --}}
{{!-- @name kicker @type outlined-textfield @default "" @label "키커(영문, 선택)" --}}
{{!-- @name headline @type text-editor @default "<p>헤드라인을 입력하세요</p>" @label "헤드라인" --}}
{{!-- @name sub @type text-editor @default "<p></p>" @label "서브 문구(선택)" --}}
{{!-- @name btnText @type outlined-textfield @default "" @label "버튼 문구(선택)" --}}
{{!-- @name btnLink @type outlined-textfield @default "" @label "버튼 링크" --}}
<div class="mm-cine" data-align="{{align}}" data-overlay="{{overlay}}" data-focus="{{focus}}" data-height="{{height}}" data-maxw="{{maxw}}">
  <img class="mm-cine__bg" src="{{image}}" alt="">
  <img class="mm-cine__bgm" src="{{imageMobile}}" alt="">
  <div class="mm-cine__veil" aria-hidden="true"></div>
  <div class="mm-cine__inner">
    <span class="mm-cine__kicker">{{kicker}}</span>
    <div class="mm-cine__headline" role="heading" aria-level="2">{{headline}}</div>
    <div class="mm-cine__sub">{{sub}}</div>
    <a class="mm-cine__btn" href="{{btnLink}}">{{btnText}}</a>
  </div>
</div>
```
### CSS 탭
```css
.mm-cine{position:relative;overflow:hidden;border-radius:{{radius}};min-height:clamp(280px,46vw,460px);display:flex;align-items:center;background:#1A1A1A;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-cine[data-align="가운데"]{justify-content:center;text-align:center;}
.mm-cine[data-align="왼쪽"]{justify-content:flex-start;text-align:left;}
.mm-cine__bg,.mm-cine__bgm{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;z-index:0;animation:mm-cine-zoom 16s ease-out both;}
.mm-cine__bgm{display:none;}
/* 모바일 전용 이미지가 있을 때(data-hasm=1)만 모바일에서 스왑 */
@media (max-width:768px){
  .mm-cine[data-hasm="1"] .mm-cine__bg{display:none;}
  .mm-cine[data-hasm="1"] .mm-cine__bgm{display:block;}
}
@keyframes mm-cine-zoom{from{transform:scale(1.0)}to{transform:scale(1.09)}}
.mm-cine__veil{position:absolute;inset:0;z-index:1;background:rgba(26,26,26,.45);}
.mm-cine__inner{position:relative;z-index:2;padding:clamp(28px,6vw,56px);max-width:720px;}
.mm-cine__kicker{display:block;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(10px,2.6vw,12px);font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#D4D0CB;margin-bottom:12px;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-cine__kicker:empty{display:none;}
.mm-cine__headline{margin:0;font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:clamp(24px,6vw,44px);font-weight:900;line-height:1.12;letter-spacing:-.02em;color:#FAF9F7;text-shadow:0 2px 24px rgba(0,0,0,.35);white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-cine__headline p{margin:0 0 .15em;}
.mm-cine__sub{margin:14px 0 0;font-size:clamp(14px,3.8vw,17px);line-height:1.45;color:#EDEBE8;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-cine__sub p{margin:0 0 .15em;}
.mm-cine__sub:empty{display:none;}
.mm-cine__btn{display:inline-flex;margin-top:24px;padding:14px 34px;border-radius:8px;background:#FAF9F7;color:#1A1A1A;font-weight:700;font-size:15px;text-decoration:none;transition:opacity .2s,transform .2s;}
.mm-cine__btn:empty{display:none;}
.mm-cine__btn:active{transform:scale(.97);}
@media (hover:hover){.mm-cine__btn:hover{opacity:.9;}}
@media (prefers-reduced-motion:reduce){.mm-cine__bg{animation:none;}}
```
### JS 탭
```js
/* 시네마틱 배너 — Ken Burns는 CSS. '어둡게'=검정+투명도, 정렬·사진초점·높이 적용. (정규식·인라인핸들러 미사용) */
(function(){
  function alphaOf(ov){
    ov=String(ov||'').trim();
    if(ov.indexOf('rgba')===0){
      var inner=ov.substring(ov.indexOf('(')+1, ov.lastIndexOf(')')),p=inner.split(',');
      if(p.length>=4){ var a=parseFloat(p[3]); if(!isNaN(a)) return a; }
      return null;
    }
    var hex=ov.charAt(0)==='#'?ov.substring(1):ov;
    if(hex.length===8){ var v=parseInt(hex.substring(6,8),16); if(!isNaN(v)) return v/255; }
    return null;
  }
  function isLeft(al){al=String(al||'').trim();var low=al.toLowerCase();
    /* 어떤 유형이 와도 인식: 텍스트(왼쪽/좌/left) · 스위치(true) · 옵션버튼(왼쪽) */
    return al==='true'||al.indexOf('왼')>=0||al.indexOf('좌')>=0||low.indexOf('left')>=0||low.indexOf('start')>=0;}
  function focusPos(f){
    f=String(f||'').trim();var low=f.toLowerCase();
    if(f.indexOf('좌')>=0||f.indexOf('왼')>=0||low.indexOf('left')>=0)return 'left center';
    if(f.indexOf('우')>=0||f.indexOf('오')>=0||low.indexOf('right')>=0)return 'right center';
    if(f.indexOf('상')>=0||f.indexOf('위')>=0||low.indexOf('top')>=0)return 'center top';
    if(f.indexOf('하')>=0||f.indexOf('아래')>=0||low.indexOf('bottom')>=0)return 'center bottom';
    return 'center';
  }
  function initOne(root){
    var ov=root.getAttribute('data-overlay'),veil=root.querySelector('.mm-cine__veil'),bg=root.querySelector('.mm-cine__bg');
    if(veil&&ov){ var a=alphaOf(ov); if(a!==null){ a=Math.max(0,Math.min(1,a)); veil.style.background='rgba(26,26,26,'+a+')'; } else { veil.style.background=ov; } }
    root.setAttribute('data-align', isLeft(root.getAttribute('data-align')) ? '왼쪽' : '가운데');
    var bgm=root.querySelector('.mm-cine__bgm'),fp=focusPos(root.getAttribute('data-focus'));
    if(bg) bg.style.objectPosition=fp;
    if(bgm) bgm.style.objectPosition=fp;
    /* 모바일 이미지가 실제 있으면 data-hasm=1 → 모바일에서 스왑. 없으면 PC 이미지 공용 */
    if(bgm && String(bgm.getAttribute('src')||'').trim()) root.setAttribute('data-hasm','1');
    var h=root.getAttribute('data-height');
    if(h){ var hv=h.trim(); if(hv){ if(String(parseFloat(hv))===hv) hv+='px'; root.style.minHeight=hv; } }
    /* 모서리(radius)는 CSS에서 {{radius}}로 직접 그림 → 로드 시 라운드→각짐 깜빡임 없음 */
    /* PC 최대 가로폭 지정 + 중앙정렬. 모바일은 화면이 더 좁아 자동 꽉 채움 */
    var mw=root.getAttribute('data-maxw');
    if(mw!==null){ var mv=mw.trim(); if(mv){ if(String(parseFloat(mv))===mv) mv+='px'; root.style.maxWidth=mv; root.style.marginLeft='auto'; root.style.marginRight='auto'; } }
  }
  function init(){var l=document.querySelectorAll('.mm-cine');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## D2. 스플릿 프로모 배너
`폴더: d02_split_promo`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "좌우 반반 이미지+카피 프로모 배너." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 스플릿 프로모 배너
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 좌우 반반 — 이미지 + 카피·CTA (신제품·이벤트 정석 배너)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name image @type image @label "PC 이미지 — 권장 900×800px" --}}
{{!-- @name imageMobile @type image @label "모바일 이미지 (선택) — 권장 1080×1080px · 비우면 PC 이미지 사용" --}}
{{!-- @name imgPos @type outlined-textfield @default "왼쪽" @label "이미지 위치 — 입력: 왼쪽 · 오른쪽" --}}
{{!-- @name theme @type outlined-textfield @default "라이트" @label "테마 — 입력: 라이트 · 다크" --}}
{{!-- @name height @type outlined-textfield @default "" @label "높이 — 예: 360px (비우면 자동)" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name tag @type outlined-textfield @default "" @label "태그칩(선택)" --}}
{{!-- @name title @type outlined-textfield @default "제목을 입력하세요" @label "제목" --}}
{{!-- @name desc @type text-editor @default "<p>설명</p>" @label "설명" --}}
{{!-- @name btnText @type outlined-textfield @default "자세히 보기" @label "버튼 문구(비우면 숨김)" --}}
{{!-- @name btnLink @type outlined-textfield @default "" @label "버튼 링크" --}}
<div class="mm-split" data-img="{{imgPos}}" data-theme="{{theme}}" data-height="{{height}}">
  <div class="mm-split__media"><img class="mm-split__img" src="{{image}}" alt=""><img class="mm-split__imgm" src="{{imageMobile}}" alt=""></div>
  <div class="mm-split__body">
    <span class="mm-split__tag">{{tag}}</span>
    <h2 class="mm-split__title">{{title}}</h2>
    <div class="mm-split__desc">{{desc}}</div>
    <a class="mm-split__btn" href="{{btnLink}}">{{btnText}}</a>
  </div>
</div>
```
### CSS 탭
```css
.mm-split{display:grid;grid-template-columns:1fr 1fr;align-items:stretch;border-radius:{{radius}};overflow:hidden;border:1px solid #D4D0CB;background:#FFFFFF;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-split[data-theme="다크"]{background:#1A1A1A;border-color:#2D2D2D;color:#FAF9F7;}
.mm-split[data-img="오른쪽"] .mm-split__media{order:2;}
.mm-split__media{min-height:240px;background:#F5F3F0;}
.mm-split__media img{width:100%;height:100%;object-fit:cover;display:block;}
.mm-split__imgm{display:none;}
.mm-split__body{padding:clamp(24px,4vw,44px);display:flex;flex-direction:column;justify-content:center;}
.mm-split__tag{display:inline-block;align-self:flex-start;font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;padding:4px 10px;border-radius:4px;background:#1A1A1A;color:#FAF9F7;margin-bottom:14px;}
.mm-split[data-theme="다크"] .mm-split__tag{background:#FAF9F7;color:#1A1A1A;}
.mm-split__tag:empty{display:none;}
.mm-split__title{margin:0 0 12px;font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:clamp(20px,4vw,30px);font-weight:800;letter-spacing:-.02em;line-height:1.25;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-split__desc{font-size:clamp(14px,3.8vw,16px);line-height:1.65;color:#4A4A4A;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-split[data-theme="다크"] .mm-split__desc{color:#D4D0CB;}
.mm-split__desc p{margin:0 0 .15em;}
.mm-split__btn{align-self:flex-start;margin-top:20px;padding:13px 30px;border-radius:8px;background:#1A1A1A;color:#FAF9F7;font-weight:700;font-size:15px;text-decoration:none;transition:opacity .2s;}
.mm-split[data-theme="다크"] .mm-split__btn{background:#FAF9F7;color:#1A1A1A;}
.mm-split__btn:empty{display:none;}
@media (hover:hover){.mm-split__btn:hover{opacity:.88;}}
@media (max-width:640px){.mm-split{grid-template-columns:1fr;}.mm-split[data-img="오른쪽"] .mm-split__media{order:0;}.mm-split[data-hasm="1"] .mm-split__img{display:none;}.mm-split[data-hasm="1"] .mm-split__imgm{display:block;}}
```
### JS 탭
```js
/* 스플릿 프로모 배너 — 레이아웃은 HTML/CSS. JS는 높이만 적용(선택). 정규식·인라인핸들러 미사용. */
(function(){
  function initOne(root){
    var h=root.getAttribute('data-height');
    if(h){ var hv=h.trim(); if(hv){ if(String(parseFloat(hv))===hv) hv+='px'; root.style.minHeight=hv; } }
    /* 모바일 이미지가 있으면 data-hasm=1 → 모바일에서 스왑 */
    var m=root.querySelector('.mm-split__imgm');
    if(m && String(m.getAttribute('src')||'').trim()) root.setAttribute('data-hasm','1');
  }
  function init(){var l=document.querySelectorAll('.mm-split');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## D3. 슬림 공지 띠
`폴더: d03_notice_bar`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "얇은 공지 띠(흐름 옵션, 상시 노출형)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 슬림 공지 띠
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 얇은 안내 띠(배송공지·이벤트). 흐름(마퀴) 옵션
  🚫 fetch·iframe 0 (※'닫기 기억'은 localStorage 금지라 상시 노출형)
═══════════════════════════════════════════════════════════════ -->
{{!-- @name icon @type outlined-textfield @default "✦" @label "아이콘(이모지/심볼, 선택)" --}}
{{!-- @name radius @type outlined-textfield @default "10px" @label "모서리 둥글기 — 예: 10px · 0px이면 각지게" --}}
{{!-- @name text @type outlined-textfield @default "안내 문구를 입력하세요" @label "문구" --}}
{{!-- @name linkText @type outlined-textfield @default "" @label "링크 문구(선택)" --}}
{{!-- @name link @type outlined-textfield @default "" @label "링크" --}}
{{!-- @name theme @type outlined-textfield @default "다크" @label "테마 — 입력: 다크 · 라이트" --}}
{{!-- @name flow @type switch @default false @label "문구 흐르기(마퀴)" --}}
<div class="mm-notice" data-theme="{{theme}}" data-flow="{{flow}}">
  <div class="mm-notice__in">
    <span class="mm-notice__icon">{{icon}}</span>
    <span class="mm-notice__text">{{text}}</span>
    <a class="mm-notice__link" href="{{link}}">{{linkText}}</a>
  </div>
</div>
```
### CSS 탭
```css
.mm-notice{width:100%;border-radius:{{radius}};background:#1A1A1A;color:#FAF9F7;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;overflow:hidden;}
.mm-notice[data-theme="라이트"]{background:#F5F3F0;color:#1A1A1A;border:1px solid #D4D0CB;}
.mm-notice__in{display:flex;align-items:center;justify-content:center;gap:10px;padding:11px 16px;font-size:clamp(13px,3.4vw,14px);font-weight:600;}
.mm-notice__icon{flex:0 0 auto;}
.mm-notice__icon:empty{display:none;}
.mm-notice__text{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.mm-notice__link{flex:0 0 auto;color:inherit;text-decoration:underline;text-underline-offset:3px;font-weight:700;}
.mm-notice__link:empty,.mm-notice__link[href=""]{display:none;}
/* 흐르기(마퀴) */
.mm-notice[data-flow="true"] .mm-notice__in{justify-content:flex-start;width:max-content;animation:mm-notice-flow 18s linear infinite;}
.mm-notice[data-flow="true"] .mm-notice__text{white-space:nowrap;overflow:visible;}
.mm-notice[data-flow="true"]:hover .mm-notice__in{animation-play-state:paused;}
@keyframes mm-notice-flow{from{transform:translateX(0)}to{transform:translateX(-50%)}}
@media (prefers-reduced-motion:reduce){.mm-notice[data-flow="true"] .mm-notice__in{animation:none;width:auto;justify-content:center;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    if(root.getAttribute('data-flow')!=='true')return;
    var inEl=root.querySelector('.mm-notice__in');
    if(!inEl)return;
    // 끊김 없는 흐름: 내용 복제(50% 지점 동일)
    inEl.innerHTML=inEl.innerHTML+'<span style="display:inline-block;width:48px"></span>'+inEl.innerHTML;
  }
  function init(){var l=document.querySelectorAll('.mm-notice');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## D4. 대형 브랜드 인용 배너
`폴더: d04_quote_banner`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "브랜드 한 문장 대형 인용 배너." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 대형 브랜드 인용 배너
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 브랜드 한 문장을 거대한 따옴표로 (선언형 배너, 여백·타이포 압도)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name theme @type outlined-textfield @default "라이트" @label "테마 — 입력: 라이트 · 다크" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name align @type outlined-textfield @default "가운데" @label "정렬 — 입력: 가운데 · 왼쪽" --}}
{{!-- @name quote @type text-editor @default "<p>좋은 가위는 권하지 않습니다</p>" @label "인용문" --}}
{{!-- @name author @type outlined-textfield @default "" @label "출처/서명(선택)" --}}
<div class="mm-quote" data-theme="{{theme}}" data-align="{{align}}">
  <span class="mm-quote__mark" aria-hidden="true">“</span>
  <blockquote class="mm-quote__text">{{quote}}</blockquote>
  <cite class="mm-quote__author">{{author}}</cite>
</div>
```
### CSS 탭
```css
.mm-quote{position:relative;max-width:860px;margin:0 auto;padding:clamp(36px,7vw,72px) clamp(20px,5vw,48px);border-radius:{{radius}};background:#FAF9F7;color:#1A1A1A;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-quote[data-theme="다크"]{background:#1A1A1A;color:#FAF9F7;}
.mm-quote[data-align="가운데"]{text-align:center;}
.mm-quote[data-align="왼쪽"]{text-align:left;}
.mm-quote__mark{display:block;font-family:'Outfit',Georgia,serif;font-size:clamp(60px,16vw,120px);line-height:.6;font-weight:900;color:#D4D0CB;margin-bottom:clamp(4px,1vw,8px);}
.mm-quote[data-theme="다크"] .mm-quote__mark{color:#4A4A4A;}
.mm-quote[data-align="가운데"] .mm-quote__mark{text-align:center;}
.mm-quote__text{margin:0;font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:clamp(22px,5.5vw,40px);font-weight:800;letter-spacing:-.02em;line-height:1.35;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-quote__text p{margin:0 0 .15em;}
.mm-quote__author{display:block;margin-top:clamp(16px,3vw,26px);font-size:clamp(12px,3vw,14px);font-weight:600;font-style:normal;letter-spacing:.04em;color:#8A8580;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-quote__author:empty{display:none;}
.mm-quote__author::before{content:'— ';}
.mm-quote__author:empty::before{content:'';}
```
### JS 탭
```js
/* 대형 인용 배너 — 정렬값 관대 인식(텍스트/세그먼트/스위치) + 진입 reveal. */
(function(){
  function isLeft(al){al=String(al||'').trim();var low=al.toLowerCase();return al==='true'||al.indexOf('왼')>=0||al.indexOf('좌')>=0||low.indexOf('left')>=0||low.indexOf('start')>=0;}
  function init(){
    var l=document.querySelectorAll('.mm-quote');
    if(!l.length){return setTimeout(init,50);}
    for(var i=0;i<l.length;i++) l[i].setAttribute('data-align', isLeft(l[i].getAttribute('data-align'))?'왼쪽':'가운데');
    if(!('IntersectionObserver' in window))return;
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.style.opacity='1';e.target.style.transform='none';io.unobserve(e.target);}});},{threshold:.3});
    for(var j=0;j<l.length;j++){l[j].style.transition='opacity .7s cubic-bezier(.4,0,.2,1),transform .7s cubic-bezier(.4,0,.2,1)';l[j].style.opacity='0';l[j].style.transform='translateY(18px)';io.observe(l[j]);}
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---
