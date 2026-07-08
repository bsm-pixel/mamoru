# 🧩 MAMORU 아임웹 커스텀 위젯 전체 모음 (집 작업용)

> 각 위젯=HTML/CSS/JS 3탭. 🚫 삼중괄호 {{{ }}} 금지 · CSS탭 {{변수}} 금지 · 인라인 `style="…{{}}…"` 금지(동적값은 `data-*` 속성 + JS로 적용) · 인라인 on*= 금지. 총 44종. 갱신 2026-07-07.


## 01. 복원 Before/After 슬라이더 ⭐
`폴더: 01_before_after`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "가위 복원 전/후를 손잡이 드래그로 비교. 전·후 사진은 같은 비율로 업로드." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 복원 Before/After 슬라이더
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 가위 복원 전/후를 손잡이 드래그로 비교
  🚫 fetch·iframe·외부라이브러리 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name beforeImage @type image @label "복원 전 사진 — 권장 1200×900px (후 사진과 동일 크기)" --}}
{{!-- @name afterImage  @type image @label "복원 후 사진 — 권장 1200×900px (전 사진과 동일 크기)" --}}
{{!-- @name ratio @type outlined-textfield @default "4/3" @label "사진 비율 — 입력: 4/3 · 1/1 · 3/4 · 16/9" --}}
{{!-- @name radius @type outlined-textfield @default "12px" @label "모서리 둥글기 — 예: 12px · 0px이면 각지게" --}}
{{!-- @name beforeLabel @type outlined-textfield @default "BEFORE" @label "왼쪽 라벨" --}}
{{!-- @name afterLabel  @type outlined-textfield @default "AFTER"  @label "오른쪽 라벨" --}}
{{!-- @name caption     @type outlined-textfield @default "" @label "하단 설명(선택)" --}}
{{!-- @name startPos    @type outlined-textfield @default "50" @label "시작 위치 %" --}}
<div class="mm-ba" data-start="{{startPos}}">
  <div class="mm-ba__stage" data-ar="{{ratio}}" data-radius="{{radius}}">
    <img class="mm-ba__img mm-ba__img--after"  src="{{afterImage}}"  alt="{{afterLabel}}"  draggable="false">
    <img class="mm-ba__img mm-ba__img--before" src="{{beforeImage}}" alt="{{beforeLabel}}" draggable="false">
    <span class="mm-ba__tag mm-ba__tag--before">{{beforeLabel}}</span>
    <span class="mm-ba__tag mm-ba__tag--after">{{afterLabel}}</span>
    <div class="mm-ba__handle" aria-hidden="true"><span class="mm-ba__grip"></span></div>
  </div>
  <p class="mm-ba__caption">{{caption}}</p>
</div>
```
### CSS 탭
```css
.mm-ba{max-width:760px;margin:0 auto;font-family:-apple-system,'Noto Sans KR',sans-serif;}
.mm-ba__stage{position:relative;width:100%;aspect-ratio:4/3;overflow:hidden;background:#F5F3F0;touch-action:none;cursor:ew-resize;user-select:none;}
.mm-ba__img{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;display:block;pointer-events:none;}
.mm-ba__img--after{z-index:1;}
.mm-ba__img--before{z-index:2;-webkit-clip-path:inset(0 50% 0 0);clip-path:inset(0 50% 0 0);}
.mm-ba__tag{position:absolute;top:12px;z-index:3;font-size:11px;font-weight:700;letter-spacing:.08em;padding:4px 10px;border-radius:999px;line-height:1.4;}
.mm-ba__tag--before{left:12px;background:rgba(26,26,26,.78);color:#FAF9F7;}
.mm-ba__tag--after{right:12px;background:#FAF9F7;color:#1A1A1A;border:1px solid #D4D0CB;}
.mm-ba__handle{position:absolute;top:0;bottom:0;left:50%;width:2px;background:#FAF9F7;transform:translateX(-1px);z-index:4;box-shadow:0 0 0 1px rgba(0,0,0,.15);pointer-events:none;}
.mm-ba__grip{position:absolute;top:50%;left:50%;width:40px;height:40px;transform:translate(-50%,-50%);background:#FAF9F7;border-radius:50%;box-shadow:0 2px 10px rgba(0,0,0,.25);}
.mm-ba__grip::before,.mm-ba__grip::after{content:'';position:absolute;top:50%;width:0;height:0;border-top:5px solid transparent;border-bottom:5px solid transparent;transform:translateY(-50%);}
.mm-ba__grip::before{left:10px;border-right:6px solid #1A1A1A;}
.mm-ba__grip::after{right:10px;border-left:6px solid #1A1A1A;}
.mm-ba__caption{text-align:center;font-size:13px;color:#8A8580;margin:12px 0 0;line-height:1.5;}
.mm-ba__caption:empty{display:none;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-ba{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var stage=root.querySelector('.mm-ba__stage'),before=root.querySelector('.mm-ba__img--before'),handle=root.querySelector('.mm-ba__handle');
    if(!stage||!before||!handle)return;
    var start=parseFloat(root.getAttribute('data-start'));if(isNaN(start))start=50;
    var pos=Math.min(100,Math.max(0,start)),dragging=false;
    function apply(p){pos=Math.min(100,Math.max(0,p));var v='inset(0 '+(100-pos)+'% 0 0)';before.style.webkitClipPath=v;before.style.clipPath=v;handle.style.left=pos+'%';}
    function fromEvent(e){var r=stage.getBoundingClientRect();var cx=(e.touches&&e.touches[0])?e.touches[0].clientX:e.clientX;apply((cx-r.left)/r.width*100);}
    stage.addEventListener('pointerdown',function(e){dragging=true;try{stage.setPointerCapture(e.pointerId);}catch(_){}fromEvent(e);});
    stage.addEventListener('pointermove',function(e){if(dragging)fromEvent(e);});
    window.addEventListener('pointerup',function(){dragging=false;});
    root.setAttribute('tabindex','0');
    root.addEventListener('keydown',function(e){if(e.key==='ArrowLeft'){apply(pos-2);e.preventDefault();}else if(e.key==='ArrowRight'){apply(pos+2);e.preventDefault();}});
    apply(pos);
  }
  function init(){var l=document.querySelectorAll('.mm-ba');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();

/* 동적 스타일(비율/배경/좌표): 인라인 style {{}}는 아임웹 저장거부 → data-* 속성을 JS로 적용 */
(function(){function ap(){
var A=document.querySelectorAll("[data-ar]");for(var i=0;i<A.length;i++){var v=(A[i].getAttribute("data-ar")||"").trim();if(v){if(A[i].classList.contains("mm-cat"))A[i].style.setProperty("--mm-ratio",v);else A[i].style.aspectRatio=v;}}
var B=document.querySelectorAll("[data-bg]");for(var i=0;i<B.length;i++){var v=(B[i].getAttribute("data-bg")||"").trim();if(v)B[i].style.backgroundImage="url('"+v+"')";}
var C=document.querySelectorAll("[data-x]");for(var i=0;i<C.length;i++){var x=(C[i].getAttribute("data-x")||"").trim(),y=(C[i].getAttribute("data-y")||"").trim();if(x)C[i].style.left=x+"%";if(y)C[i].style.top=y+"%";}
}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",ap);else ap();})();
```

---

## 02. 한정세일 카운트다운
`폴더: 02_countdown`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "종료 일시까지 실시간 카운트다운. 0이 되면 종료 문구로 전환." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 한정세일 카운트다운
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 종료 일시까지 실시간 카운트다운. 0이 되면 종료 문구로 전환.
  🚫 fetch·iframe·외부라이브러리·localStorage 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "한정 혜택 종료까지" @label "제목" --}}
{{!-- @name endDate @type date-picker @label "종료 날짜" --}}
{{!-- @name endTime @type time-picker @label "종료 시간" --}}
{{!-- @name theme @type outlined-textfield @default "라이트" @label "테마 — 입력: 라이트 · 다크" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name btnText @type outlined-textfield @default "지금 보러가기" @label "버튼 문구(비우면 숨김)" --}}
{{!-- @name btnLink @type outlined-textfield @default "" @label "버튼 링크" --}}
{{!-- @name endedText @type outlined-textfield @default "이번 혜택은 마감되었습니다" @label "종료 후 문구" --}}
<div class="mm-cd" data-radius="{{radius}}" data-theme="{{theme}}" data-date="{{endDate}}" data-time="{{endTime}}">
  <p class="mm-cd__title">{{title}}</p>
  <div class="mm-cd__timer" aria-live="off">
    <div class="mm-cd__unit"><span class="mm-cd__num" data-cd="d">00</span><span class="mm-cd__lbl">일</span></div>
    <span class="mm-cd__sep">:</span>
    <div class="mm-cd__unit"><span class="mm-cd__num" data-cd="h">00</span><span class="mm-cd__lbl">시</span></div>
    <span class="mm-cd__sep">:</span>
    <div class="mm-cd__unit"><span class="mm-cd__num" data-cd="m">00</span><span class="mm-cd__lbl">분</span></div>
    <span class="mm-cd__sep">:</span>
    <div class="mm-cd__unit"><span class="mm-cd__num" data-cd="s">00</span><span class="mm-cd__lbl">초</span></div>
  </div>
  <a class="mm-cd__btn" href="{{btnLink}}">{{btnText}}</a>
  <p class="mm-cd__ended" data-ended="{{endedText}}"></p>
</div>
```
### CSS 탭
```css
.mm-cd{max-width:520px;margin:0 auto;padding:clamp(24px,5vw,36px) clamp(20px,4vw,32px);text-align:center;
  font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;
  background:#FAF9F7;border:1px solid #D4D0CB;color:#1A1A1A;}
.mm-cd[data-theme="다크"]{background:#1A1A1A;border-color:#2D2D2D;color:#FAF9F7;}
.mm-cd__title{margin:0 0 18px;font-size:clamp(13px,3.6vw,15px);font-weight:600;letter-spacing:.02em;color:#8A8580;}
.mm-cd[data-theme="다크"] .mm-cd__title{color:#B8B4AF;}
.mm-cd__timer{display:flex;align-items:flex-start;justify-content:center;gap:clamp(6px,2vw,12px);}
.mm-cd__unit{display:flex;flex-direction:column;align-items:center;min-width:clamp(48px,15vw,64px);}
.mm-cd__num{font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(30px,9vw,48px);font-weight:800;line-height:1;letter-spacing:-.02em;font-variant-numeric:tabular-nums;}
.mm-cd__lbl{margin-top:8px;font-size:11px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#8A8580;}
.mm-cd[data-theme="다크"] .mm-cd__lbl{color:#8A8580;}
.mm-cd__sep{font-family:'Outfit',sans-serif;font-size:clamp(24px,7vw,38px);font-weight:300;line-height:1;color:#D4D0CB;padding-top:4px;}
.mm-cd[data-theme="다크"] .mm-cd__sep{color:#4A4A4A;}
.mm-cd__btn{display:inline-flex;align-items:center;justify-content:center;margin-top:24px;padding:14px 32px;border-radius:8px;font-size:15px;font-weight:700;text-decoration:none;transition:transform .2s cubic-bezier(.4,0,.2,1),opacity .2s;
  background:#1A1A1A;color:#FAF9F7;}
.mm-cd[data-theme="다크"] .mm-cd__btn{background:#FAF9F7;color:#1A1A1A;}
.mm-cd__btn:empty{display:none;}
.mm-cd__btn:active{transform:scale(.97);}
@media (hover:hover){.mm-cd__btn:hover{opacity:.88;}}
.mm-cd__ended{display:none;margin:0;font-size:clamp(16px,4.5vw,20px);font-weight:700;letter-spacing:-.01em;}
.mm-cd.is-ended .mm-cd__timer,.mm-cd.is-ended .mm-cd__btn,.mm-cd.is-ended .mm-cd__title{display:none;}
.mm-cd.is-ended .mm-cd__ended{display:block;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-cd{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function parseTarget(dateStr,timeStr){
    if(!dateStr) return null;
    var dm=String(dateStr).match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
    if(!dm) return null;
    var y=+dm[1],mo=+dm[2]-1,d=+dm[3],h=23,mi=59;
    if(timeStr){
      var pm=/오후|PM/i.test(timeStr),am=/오전|AM/i.test(timeStr);
      var tm=String(timeStr).match(/(\d{1,2})\D+(\d{2})/);
      if(tm){h=+tm[1];mi=+tm[2];if(pm&&h<12)h+=12;if(am&&h===12)h=0;}
    }
    var t=new Date(y,mo,d,h,mi,0).getTime();
    return isNaN(t)?null:t;
  }
  function pad(n){return n<10?'0'+n:''+n;}
  function initOne(root){
    var target=parseTarget(root.getAttribute('data-date'),root.getAttribute('data-time'));
    var endedEl=root.querySelector('.mm-cd__ended');
    var endedText=endedEl?endedEl.getAttribute('data-ended'):'';
    if(endedEl&&endedText) endedEl.textContent=endedText;
    var nums={d:root.querySelector('[data-cd="d"]'),h:root.querySelector('[data-cd="h"]'),m:root.querySelector('[data-cd="m"]'),s:root.querySelector('[data-cd="s"]')};
    if(!target){root.classList.add('is-ended');return;}
    var timer=null;
    function tick(){
      var diff=target-Date.now();
      if(diff<=0){root.classList.add('is-ended');if(timer)clearInterval(timer);return;}
      var s=Math.floor(diff/1000);
      var d=Math.floor(s/86400);s-=d*86400;
      var h=Math.floor(s/3600);s-=h*3600;
      var m=Math.floor(s/60);s-=m*60;
      if(nums.d)nums.d.textContent=pad(d);
      if(nums.h)nums.h.textContent=pad(h);
      if(nums.m)nums.m.textContent=pad(m);
      if(nums.s)nums.s.textContent=pad(s);
    }
    tick();timer=setInterval(tick,1000);
  }
  function init(){var l=document.querySelectorAll('.mm-cd');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
```

---

## 03. 가위 스펙 비교표
`폴더: 03_compare`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "모델 2~3개 스펙을 나란히 비교(모바일 가로스크롤)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 가위 스펙 비교표
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 모델 2~3개 나란히 스펙 비교(모바일 가로스크롤). 구매 결정 가속.
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name models @type item @label "모델" --}}
<div class="mm-cmp">
  <div class="mm-cmp__row">
  {{#each models}}
    {{!-- @name name @type outlined-textfield @default "모델명" @label "모델명" --}}
    {{!-- @name image @type image @label "이미지(선택) (권장 800×800px)" --}}
    {{!-- @name weight @type outlined-textfield @default "" @label "무게" --}}
    {{!-- @name length @type outlined-textfield @default "" @label "길이" --}}
    {{!-- @name steel @type outlined-textfield @default "" @label "강재" --}}
    {{!-- @name use @type outlined-textfield @default "" @label "용도" --}}
    {{!-- @name link @type outlined-textfield @default "" @label "상품 링크(선택)" --}}
    <article class="mm-cmp__card">
      <img class="mm-cmp__img" src="{{image}}" alt="">
      <h3 class="mm-cmp__name">{{name}}</h3>
      <dl class="mm-cmp__specs">
        <div class="mm-cmp__spec"><dt>무게</dt><dd>{{weight}}</dd></div>
        <div class="mm-cmp__spec"><dt>길이</dt><dd>{{length}}</dd></div>
        <div class="mm-cmp__spec"><dt>강재</dt><dd>{{steel}}</dd></div>
        <div class="mm-cmp__spec"><dt>용도</dt><dd>{{use}}</dd></div>
      </dl>
      <a class="mm-cmp__link" href="{{link}}">자세히 보기</a>
    </article>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-cmp{max-width:880px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-cmp__row{display:grid;grid-template-columns:repeat(auto-fit,minmax(0,1fr));gap:14px;}
@media (max-width:600px){.mm-cmp__row{display:flex;overflow-x:auto;scroll-snap-type:x mandatory;-webkit-overflow-scrolling:touch;gap:12px;padding-bottom:6px;}.mm-cmp__row::-webkit-scrollbar{display:none;}.mm-cmp__card{flex:0 0 78%;scroll-snap-align:start;}}
.mm-cmp__card{background:#FFFFFF;border:1px solid #D4D0CB;border-radius:14px;padding:18px;display:flex;flex-direction:column;}
.mm-cmp__img{width:100%;aspect-ratio:4/3;object-fit:cover;border-radius:10px;background:#F5F3F0;margin-bottom:14px;display:block;}
.mm-cmp__name{margin:0 0 14px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(16px,4vw,19px);font-weight:800;letter-spacing:-.01em;}
.mm-cmp__specs{margin:0;display:flex;flex-direction:column;gap:0;flex:1;}
.mm-cmp__spec{display:flex;justify-content:space-between;gap:12px;padding:10px 0;border-top:1px solid #EDEBE8;}
.mm-cmp__spec dt{margin:0;font-size:12px;font-weight:600;color:#8A8580;letter-spacing:.02em;}
.mm-cmp__spec dd{margin:0;font-size:clamp(13px,3.6vw,15px);font-weight:600;text-align:right;}
.mm-cmp__spec dd:empty::after{content:'—';color:#B8B4AF;}
.mm-cmp__link{margin-top:16px;text-align:center;padding:11px;border-radius:8px;background:#1A1A1A;color:#FAF9F7;font-size:14px;font-weight:700;text-decoration:none;transition:opacity .2s;}
.mm-cmp__link[href=""],.mm-cmp__link:not([href]){display:none;}
@media (hover:hover){.mm-cmp__link:hover{opacity:.88;}}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-cmp{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
/* 스펙 비교표 — 카드 진입 reveal(선택). 핵심 기능은 HTML/CSS만으로 동작. */
(function(){
  function initOne(root){
    if(!('IntersectionObserver' in window))return;
    var cards=root.querySelectorAll('.mm-cmp__card');
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.style.opacity='1';e.target.style.transform='none';io.unobserve(e.target);}});},{threshold:.15});
    for(var i=0;i<cards.length;i++){cards[i].style.transition='opacity .5s cubic-bezier(.4,0,.2,1) '+(i*0.06)+'s, transform .5s cubic-bezier(.4,0,.2,1) '+(i*0.06)+'s';cards[i].style.opacity='0';cards[i].style.transform='translateY(18px)';io.observe(cards[i]);}
  }
  function init(){var l=document.querySelectorAll('.mm-cmp');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## 04. 등급별 견적 계산기
`폴더: 04_calculator`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "모델·수량·등급 선택 시 예상 견적 자동 계산." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 등급별 견적 계산기
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 모델·수량·등급 선택 → 예상금액 즉시 (B2B 리드 질↑)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "예상 견적 계산" @label "제목" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name g1n @type outlined-textfield @default "소비자" @label "등급1 이름" --}}
{{!-- @name g1d @type outlined-textfield @default "0" @label "등급1 할인%" --}}
{{!-- @name g2n @type outlined-textfield @default "딜러" @label "등급2 이름" --}}
{{!-- @name g2d @type outlined-textfield @default "15" @label "등급2 할인%" --}}
{{!-- @name g3n @type outlined-textfield @default "아카데미" @label "등급3 이름" --}}
{{!-- @name g3d @type outlined-textfield @default "20" @label "등급3 할인%" --}}
{{!-- @name note @type outlined-textfield @default "실제 가격은 상담 시 확정됩니다" @label "하단 안내(선택)" --}}
{{!-- @name models @type item @label "모델·단가" --}}
<div class="mm-calc" data-radius="{{radius}}" data-g1n="{{g1n}}" data-g1d="{{g1d}}" data-g2n="{{g2n}}" data-g2d="{{g2d}}" data-g3n="{{g3n}}" data-g3d="{{g3d}}">
  <p class="mm-calc__title">{{title}}</p>
  <div class="mm-calc__data" hidden>
  {{#each models}}
    {{!-- @name mName @type outlined-textfield @default "모델명" @label "모델명" --}}
    {{!-- @name mPrice @type outlined-textfield @default "0" @label "단가(숫자만)" --}}
    <span data-name="{{mName}}" data-price="{{mPrice}}"></span>
  {{/each}}
  </div>
  <label class="mm-calc__field"><span>모델</span><select class="mm-calc__model"></select></label>
  <label class="mm-calc__field"><span>수량</span><input type="number" class="mm-calc__qty" value="1" min="1" inputmode="numeric"></label>
  <div class="mm-calc__field"><span>등급</span><div class="mm-calc__grades"></div></div>
  <div class="mm-calc__out"><span>예상 금액</span><strong class="mm-calc__total">—</strong></div>
  <p class="mm-calc__note">{{note}}</p>
</div>
```
### CSS 탭
```css
.mm-calc{max-width:460px;margin:0 auto;padding:clamp(22px,5vw,32px);border:1px solid #D4D0CB;background:#FAF9F7;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-calc__title{margin:0 0 20px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(18px,4.5vw,22px);font-weight:800;letter-spacing:-.02em;text-align:center;}
.mm-calc__field{display:flex;align-items:center;gap:14px;margin-bottom:14px;}
.mm-calc__field>span{flex:0 0 48px;font-size:13px;font-weight:600;color:#8A8580;}
.mm-calc__model,.mm-calc__qty{flex:1;padding:12px 14px;border:1px solid #D4D0CB;border-radius:8px;background:#FFFFFF;font-family:inherit;font-size:15px;color:#1A1A1A;-webkit-appearance:none;appearance:none;}
.mm-calc__grades{flex:1;display:flex;gap:6px;}
.mm-calc__gbtn{flex:1;padding:11px 4px;border:1px solid #D4D0CB;border-radius:8px;background:#FFFFFF;font-family:inherit;font-size:13px;font-weight:600;color:#1A1A1A;cursor:pointer;transition:all .2s;}
.mm-calc__gbtn.is-on{background:#1A1A1A;color:#FAF9F7;border-color:#1A1A1A;}
.mm-calc__out{display:flex;align-items:baseline;justify-content:space-between;margin-top:20px;padding-top:18px;border-top:1px solid #D4D0CB;}
.mm-calc__out>span{font-size:14px;font-weight:600;color:#8A8580;}
.mm-calc__total{font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(24px,7vw,34px);font-weight:900;letter-spacing:-.02em;font-variant-numeric:tabular-nums;}
.mm-calc__note{margin:14px 0 0;font-size:12px;color:#B8B4AF;text-align:center;}
.mm-calc__note:empty{display:none;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-calc{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function num(v){return parseFloat(String(v).replace(/[^0-9.]/g,''))||0;}
  function fmt(n){return Math.round(n).toLocaleString('en-US');}
  function initOne(root){
    var spans=root.querySelectorAll('.mm-calc__data span');
    var sel=root.querySelector('.mm-calc__model');
    var qty=root.querySelector('.mm-calc__qty');
    var gradesWrap=root.querySelector('.mm-calc__grades');
    var totalEl=root.querySelector('.mm-calc__total');
    if(!sel||!qty||!totalEl)return;
    // 모델 옵션
    var models=[];
    for(var i=0;i<spans.length;i++){var nm=spans[i].getAttribute('data-name')||('모델'+(i+1));var pr=num(spans[i].getAttribute('data-price'));models.push({name:nm,price:pr});var o=document.createElement('option');o.value=i;o.textContent=nm;sel.appendChild(o);}
    // 등급
    var grades=[];
    for(var g=1;g<=3;g++){var gn=root.getAttribute('data-g'+g+'n');if(gn){grades.push({name:gn,disc:num(root.getAttribute('data-g'+g+'d'))});}}
    var gi=0,gbtns=[];
    for(var k=0;k<grades.length;k++){(function(idx){var b=document.createElement('button');b.type='button';b.className='mm-calc__gbtn'+(idx===0?' is-on':'');b.textContent=grades[idx].name;b.addEventListener('click',function(){gi=idx;for(var x=0;x<gbtns.length;x++)gbtns[x].classList.toggle('is-on',x===idx);calc();});gradesWrap.appendChild(b);gbtns.push(b);})(k);}
    function calc(){
      var m=models[sel.value|0]||{price:0};
      var q=Math.max(1,num(qty.value));
      var disc=grades[gi]?grades[gi].disc:0;
      var total=m.price*q*(1-disc/100);
      totalEl.textContent=fmt(total)+'원';
    }
    sel.addEventListener('change',calc);qty.addEventListener('input',calc);
    calc();
  }
  function init(){var l=document.querySelectorAll('.mm-calc');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
```

---

## 05. 정적 누적 카운터
`폴더: 05_counter`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "스크롤 시 0→목표 숫자 카운트업(패널 입력값, 실시간 아님)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 정적 누적 카운터
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 스크롤 시 0→목표 숫자로 카운트업 (신뢰·규모감)
  🚫 fetch·iframe·외부라이브러리 0 (※실시간 DB 아님 — 패널에 숫자 입력)
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "" @label "제목(선택)" --}}
{{!-- @name stats @type item @label "수치 항목" --}}
<div class="mm-stat">
  <p class="mm-stat__title">{{title}}</p>
  <div class="mm-stat__grid">
  {{#each stats}}
    {{!-- @name value @type outlined-textfield @default "0" @label "숫자(예: 1240)" --}}
    {{!-- @name prefix @type outlined-textfield @default "" @label "앞 기호(선택)" --}}
    {{!-- @name suffix @type outlined-textfield @default "" @label "단위(건/% 등)" --}}
    {{!-- @name label @type outlined-textfield @default "항목" @label "라벨" --}}
    <div class="mm-stat__item">
      <span class="mm-stat__num" data-target="{{value}}" data-prefix="{{prefix}}" data-suffix="{{suffix}}">0</span>
      <span class="mm-stat__label">{{label}}</span>
    </div>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-stat{max-width:880px;margin:0 auto;padding:clamp(28px,5vw,44px) 20px;text-align:center;
  font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-stat__title{margin:0 0 clamp(20px,4vw,32px);font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(20px,5vw,28px);font-weight:800;letter-spacing:-.02em;}
.mm-stat__title:empty{display:none;}
.mm-stat__grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:clamp(20px,4vw,40px);}
.mm-stat__item{display:flex;flex-direction:column;align-items:center;gap:10px;}
.mm-stat__num{font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(38px,11vw,64px);font-weight:900;line-height:1;letter-spacing:-.03em;font-variant-numeric:tabular-nums;}
.mm-stat__label{font-size:clamp(13px,3.4vw,15px);font-weight:600;color:#8A8580;letter-spacing:.01em;}
@media (prefers-reduced-motion:no-preference){
  .mm-stat__item{opacity:0;transform:translateY(20px);transition:opacity .6s cubic-bezier(.4,0,.2,1),transform .6s cubic-bezier(.4,0,.2,1);}
  .mm-stat.is-in .mm-stat__item{opacity:1;transform:none;}
  .mm-stat.is-in .mm-stat__item:nth-child(2){transition-delay:.08s;}
  .mm-stat.is-in .mm-stat__item:nth-child(3){transition-delay:.16s;}
  .mm-stat.is-in .mm-stat__item:nth-child(4){transition-delay:.24s;}
}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-stat{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function fmt(n){return n.toLocaleString('en-US');}
  function run(root){
    var nums=root.querySelectorAll('.mm-stat__num');
    for(var i=0;i<nums.length;i++)(function(el){
      var target=parseFloat(String(el.getAttribute('data-target')).replace(/[^0-9.]/g,''))||0;
      var pre=el.getAttribute('data-prefix')||'',suf=el.getAttribute('data-suffix')||'';
      var dur=1400,t0=null;
      function step(ts){if(!t0)t0=ts;var p=Math.min((ts-t0)/dur,1);var eased=1-Math.pow(1-p,3);
        el.textContent=pre+fmt(Math.round(target*eased))+suf;
        if(p<1)requestAnimationFrame(step);}
      requestAnimationFrame(step);
    })(nums[i]);
  }
  function initOne(root){
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){root.classList.add('is-in');run(root);io.disconnect();}});},{threshold:.3});
      io.observe(root);
    }else{root.classList.add('is-in');run(root);}
  }
  function init(){var l=document.querySelectorAll('.mm-stat');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## 06. 후기 캐러셀 (수동입력)
`폴더: 06_review_carousel`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "사진·별점·후기 카드 자동 슬라이드(후기 직접 입력)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 후기 캐러셀 (수동입력)
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 사진+별점+한줄 후기 카드 자동 슬라이드 (사회적 증거)
  🚫 fetch·iframe 0 (※TMS 자동연동 아님 — 패널에 후기 직접 입력)
═══════════════════════════════════════════════════════════════ -->
{{!-- @name speed @type outlined-textfield @default "4" @label "자동전환 간격(초)" --}}
{{!-- @name reviews @type item @label "후기" --}}
<div class="mm-rev" data-speed="{{speed}}">
  <div class="mm-rev__viewport">
    <div class="mm-rev__track">
    {{#each reviews}}
      {{!-- @name photo @type image @label "사진(선택) (권장 800×800px)" --}}
      {{!-- @name rating @type outlined-textfield @default "5" @label "별점 — 입력: 5 · 4 · 3 · 2 · 1" --}}
      {{!-- @name text @type text-editor @default "<p>후기 내용을 입력하세요</p>" @label "후기 내용" --}}
      {{!-- @name name @type outlined-textfield @default "" @label "이름/매장(선택)" --}}
      <article class="mm-rev__card">
        <img class="mm-rev__photo" src="{{photo}}" alt="">
        <div class="mm-rev__stars" data-rating="{{rating}}" aria-hidden="true"></div>
        <div class="mm-rev__text">{{text}}</div>
        <p class="mm-rev__name">{{name}}</p>
      </article>
    {{/each}}
    </div>
  </div>
  <div class="mm-rev__dots" aria-hidden="true"></div>
</div>
```
### CSS 탭
```css
.mm-rev{max-width:560px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-rev__viewport{overflow:hidden;border-radius:16px;}
.mm-rev__track{display:flex;transition:transform .5s cubic-bezier(.4,0,.2,1);}
.mm-rev__card{flex:0 0 100%;box-sizing:border-box;padding:clamp(24px,5vw,36px);background:#FFFFFF;border:1px solid #D4D0CB;border-radius:16px;text-align:center;}
.mm-rev__photo{width:100%;aspect-ratio:16/10;object-fit:cover;border-radius:10px;background:#F5F3F0;margin-bottom:18px;display:block;}
.mm-rev__stars{font-size:16px;letter-spacing:2px;color:#1A1A1A;margin-bottom:14px;min-height:18px;}
.mm-rev__text{font-size:clamp(15px,4vw,17px);line-height:1.65;color:#2D2D2D;}
.mm-rev__text p{margin:0 0 .15em;}
.mm-rev__name{margin:16px 0 0;font-size:13px;font-weight:700;color:#8A8580;}
.mm-rev__name:empty{display:none;}
.mm-rev__dots{display:flex;gap:7px;justify-content:center;margin-top:16px;}
.mm-rev__dot{width:7px;height:7px;border-radius:50%;border:none;padding:0;background:#D4D0CB;cursor:pointer;transition:all .25s;}
.mm-rev__dot.is-on{background:#1A1A1A;width:22px;border-radius:4px;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-rev{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function stars(n){n=Math.max(0,Math.min(5,parseInt(n,10)||0));var s='';for(var i=0;i<5;i++)s+=(i<n?'★':'☆');return s;}
  function initOne(root){
    var track=root.querySelector('.mm-rev__track');
    var cards=root.querySelectorAll('.mm-rev__card');
    var dotsWrap=root.querySelector('.mm-rev__dots');
    var total=cards.length;
    if(!track||!total)return;
    // 별점 렌더
    for(var i=0;i<total;i++){var st=cards[i].querySelector('.mm-rev__stars');if(st)st.textContent=stars(st.getAttribute('data-rating'));}
    var cur=0,timer=null;
    var speed=(parseFloat(root.getAttribute('data-speed'))||4)*1000;
    // 점
    var dots=[];
    if(dotsWrap&&total>1){for(var d=0;d<total;d++){(function(idx){var b=document.createElement('button');b.className='mm-rev__dot'+(idx===0?' is-on':'');b.setAttribute('aria-label',(idx+1)+'번 후기');b.addEventListener('click',function(){go(idx);restart();});dotsWrap.appendChild(b);dots.push(b);})(d);}}
    function go(n){cur=(n%total+total)%total;track.style.transform='translateX(-'+(cur*100)+'%)';for(var k=0;k<dots.length;k++)dots[k].classList.toggle('is-on',k===cur);}
    function next(){go(cur+1);}
    function start(){if(total>1&&speed>0){stop();timer=setInterval(next,speed);}}
    function stop(){if(timer){clearInterval(timer);timer=null;}}
    function restart(){start();}
    root.addEventListener('mouseenter',stop);root.addEventListener('mouseleave',start);
    // 스와이프
    var sx=0;
    track.addEventListener('touchstart',function(e){sx=e.changedTouches[0].clientX;stop();},{passive:true});
    track.addEventListener('touchend',function(e){var dx=e.changedTouches[0].clientX-sx;if(Math.abs(dx)>40){dx<0?next():go(cur-1);}start();},{passive:true});
    go(0);start();
  }
  function init(){var l=document.querySelectorAll('.mm-rev');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## 07. 가위 관리법 가이드 (아코디언)
`폴더: 07_care_guide`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "관리법 단계 아코디언(한 번에 하나 펼침)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 가위 관리법 가이드(아코디언)
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 소독·보관·텐션 등 단계별 아코디언 (전문성·체류↑)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "" @label "제목(선택)" --}}
{{!-- @name items @type item @label "항목" --}}
<div class="mm-acc">
  <p class="mm-acc__title">{{title}}</p>
  {{#each items}}
    {{!-- @name q @type outlined-textfield @default "항목 제목" @label "제목" --}}
    {{!-- @name a @type text-editor @default "<p>내용</p>" @label "내용" --}}
    {{!-- @name img @type image @label "이미지(선택) (권장 800×800px)" --}}
    <div class="mm-acc__item">
      <button type="button" class="mm-acc__head"><span>{{q}}</span><i class="mm-acc__chev" aria-hidden="true"></i></button>
      <div class="mm-acc__panel"><div class="mm-acc__inner">
        <img class="mm-acc__img" src="{{img}}" alt="">
        <div class="mm-acc__a">{{a}}</div>
      </div></div>
    </div>
  {{/each}}
</div>
```
### CSS 탭
```css
.mm-acc{max-width:680px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-acc__title{margin:0 0 16px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(20px,5vw,28px);font-weight:800;letter-spacing:-.02em;}
.mm-acc__title:empty{display:none;}
.mm-acc__item{border-top:1px solid #D4D0CB;}
.mm-acc__item:last-child{border-bottom:1px solid #D4D0CB;}
.mm-acc__head{width:100%;display:flex;align-items:center;justify-content:space-between;gap:16px;padding:18px 4px;background:none;border:none;cursor:pointer;font-family:inherit;font-size:clamp(15px,4vw,17px);font-weight:700;color:#1A1A1A;text-align:left;}
.mm-acc__chev{flex:0 0 auto;width:11px;height:11px;border-right:2px solid #8A8580;border-bottom:2px solid #8A8580;transform:rotate(45deg);transition:transform .3s cubic-bezier(.4,0,.2,1);margin-right:4px;}
.mm-acc__item.is-open .mm-acc__chev{transform:rotate(-135deg);}
.mm-acc__panel{overflow:hidden;max-height:0;transition:max-height .35s cubic-bezier(.4,0,.2,1);}
.mm-acc__inner{padding:0 4px 20px;}
.mm-acc__img{width:100%;max-width:360px;aspect-ratio:16/10;object-fit:cover;border-radius:10px;background:#F5F3F0;margin-bottom:12px;display:block;}
.mm-acc__a{font-size:clamp(14px,3.8vw,15px);line-height:1.7;color:#4A4A4A;}
.mm-acc__a p{margin:0 0 .15em;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-acc{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
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
```

---

## 08. 시술별 가위 선택 가이드 (필터)
`폴더: 08_treatment_filter`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "시술 태그로 적합 모델만 필터링." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 시술별 가위 선택 가이드(필터)
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 시술(태그) 선택 → 적합 모델만 필터링 (진입장벽↓)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name allLabel @type outlined-textfield @default "전체" @label "'전체' 버튼 문구" --}}
{{!-- @name models @type item @label "모델" --}}
<div class="mm-flt" data-all="{{allLabel}}">
  <div class="mm-flt__chips"></div>
  <div class="mm-flt__grid">
  {{#each models}}
    {{!-- @name name @type outlined-textfield @default "모델명" @label "모델명" --}}
    {{!-- @name image @type image @label "이미지(선택) (권장 800×800px)" --}}
    {{!-- @name tags @type outlined-textfield @default "" @label "시술 태그(쉼표로: 커트,틴닝)" --}}
    {{!-- @name link @type outlined-textfield @default "" @label "상품 링크(선택)" --}}
    <a class="mm-flt__card" data-tags="{{tags}}" href="{{link}}">
      <img class="mm-flt__img" src="{{image}}" alt="">
      <h3 class="mm-flt__name">{{name}}</h3>
      <p class="mm-flt__tags">{{tags}}</p>
    </a>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-flt{max-width:880px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-flt__chips{display:flex;flex-wrap:wrap;gap:8px;justify-content:center;margin-bottom:22px;}
.mm-flt__chip{appearance:none;cursor:pointer;padding:9px 18px;border-radius:999px;border:1px solid #D4D0CB;background:#FFFFFF;color:#1A1A1A;font-family:inherit;font-size:14px;font-weight:600;transition:all .2s;}
.mm-flt__chip.is-on{background:#1A1A1A;color:#FAF9F7;border-color:#1A1A1A;}
@media (hover:hover){.mm-flt__chip:hover{border-color:#1A1A1A;}}
.mm-flt__grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:14px;}
.mm-flt__card{display:flex;flex-direction:column;background:#FFFFFF;border:1px solid #D4D0CB;border-radius:14px;padding:14px;text-decoration:none;color:inherit;transition:opacity .3s,transform .3s,box-shadow .25s;}
.mm-flt__card.is-hide{display:none;}
@media (hover:hover){.mm-flt__card:hover{box-shadow:0 6px 20px rgba(0,0,0,.07);}}
.mm-flt__img{width:100%;aspect-ratio:4/3;object-fit:cover;border-radius:10px;background:#F5F3F0;margin-bottom:12px;display:block;}
.mm-flt__name{margin:0 0 6px;font-size:clamp(15px,4vw,17px);font-weight:800;letter-spacing:-.01em;}
.mm-flt__tags{margin:0;font-size:12px;color:#8A8580;font-weight:600;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-flt{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function splitTags(s){return String(s||'').split(',').map(function(t){return t.trim();}).filter(Boolean);}
  function initOne(root){
    var chipsWrap=root.querySelector('.mm-flt__chips');
    var cards=root.querySelectorAll('.mm-flt__card');
    var allLabel=root.getAttribute('data-all')||'전체';
    if(!chipsWrap||!cards.length)return;
    // 태그 수집
    var set={},order=[];
    for(var i=0;i<cards.length;i++){var ts=splitTags(cards[i].getAttribute('data-tags'));for(var j=0;j<ts.length;j++){if(!set[ts[j]]){set[ts[j]]=1;order.push(ts[j]);}}}
    var chips=[];
    function mkChip(label,tag,on){var b=document.createElement('button');b.type='button';b.className='mm-flt__chip'+(on?' is-on':'');b.textContent=label;b.addEventListener('click',function(){select(tag);});chipsWrap.appendChild(b);chips.push({el:b,tag:tag});}
    mkChip(allLabel,'__all',true);
    for(var k=0;k<order.length;k++)mkChip(order[k],order[k],false);
    function select(tag){
      for(var c=0;c<chips.length;c++)chips[c].el.classList.toggle('is-on',chips[c].tag===tag);
      for(var m=0;m<cards.length;m++){var has=tag==='__all'||splitTags(cards[m].getAttribute('data-tags')).indexOf(tag)>=0;cards[m].classList.toggle('is-hide',!has);}
    }
    if(order.length===0)chipsWrap.style.display='none';
  }
  function init(){var l=document.querySelectorAll('.mm-flt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## 09. 미용가위 용어사전 (검색)
`폴더: 09_glossary`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "용어 검색/필터로 뜻 표시." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 미용가위 용어사전(검색)
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 용어 검색/필터 → 뜻 표시 (교육·SEO·전문성)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name placeholder @type outlined-textfield @default "용어 검색 (예: 텐션)" @label "검색창 안내문구" --}}
{{!-- @name terms @type item @label "용어" --}}
<div class="mm-gl">
  <input type="search" class="mm-gl__search" placeholder="{{placeholder}}" aria-label="용어 검색">
  <ul class="mm-gl__list">
  {{#each terms}}
    {{!-- @name term @type outlined-textfield @default "용어" @label "용어" --}}
    {{!-- @name def @type text-editor @default "<p>뜻</p>" @label "뜻/설명" --}}
    <li class="mm-gl__item" data-term="{{term}}">
      <dt class="mm-gl__term">{{term}}</dt>
      <dd class="mm-gl__def">{{def}}</dd>
    </li>
  {{/each}}
  </ul>
  <p class="mm-gl__empty" hidden>검색 결과가 없어요</p>
</div>
```
### CSS 탭
```css
.mm-gl{max-width:640px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-gl__search{width:100%;box-sizing:border-box;padding:14px 18px;border:1px solid #D4D0CB;border-radius:10px;background:#FFFFFF;font-family:inherit;font-size:15px;color:#1A1A1A;margin-bottom:8px;-webkit-appearance:none;}
.mm-gl__search:focus{outline:none;border-color:#1A1A1A;}
.mm-gl__list{list-style:none;margin:0;padding:0;}
.mm-gl__item{padding:18px 4px;border-bottom:1px solid #EDEBE8;}
.mm-gl__item.is-hide{display:none;}
.mm-gl__term{margin:0 0 6px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(16px,4vw,18px);font-weight:800;letter-spacing:-.01em;}
.mm-gl__def{margin:0;font-size:clamp(14px,3.8vw,15px);line-height:1.65;color:#4A4A4A;}
.mm-gl__def p{margin:0 0 .15em;}
.mm-gl__empty{text-align:center;color:#8A8580;font-size:14px;padding:24px 0;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-gl{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var search=root.querySelector('.mm-gl__search');
    var items=root.querySelectorAll('.mm-gl__item');
    var empty=root.querySelector('.mm-gl__empty');
    if(!search)return;
    function txt(el){return (el.textContent||'').toLowerCase();}
    search.addEventListener('input',function(){
      var q=search.value.trim().toLowerCase(),shown=0;
      for(var i=0;i<items.length;i++){var hit=!q||txt(items[i]).indexOf(q)>=0;items[i].classList.toggle('is-hide',!hit);if(hit)shown++;}
      if(empty)empty.hidden=shown>0;
    });
  }
  function init(){var l=document.querySelectorAll('.mm-gl');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## 10. 영업시간 배지 + 카톡 버튼
`폴더: 10_hours_badge`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "현재 시각 기준 영업중/외 자동 표시 + 카톡 버튼." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 영업시간 배지 + 카톡 버튼
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 현재 시각 기준 영업 중/외 자동 판별 + 카톡 상담 버튼
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name openLabel @type outlined-textfield @default "지금 영업 중" @label "영업중 문구" --}}
{{!-- @name closeLabel @type outlined-textfield @default "영업 종료" @label "영업외 문구" --}}
{{!-- @name btnText @type outlined-textfield @default "카카오톡 상담" @label "버튼 문구" --}}
{{!-- @name btnLink @type outlined-textfield @default "" @label "카톡 채널 링크" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name hours @type item @label "요일별 영업시간" --}}
<div class="mm-hr" data-radius="{{radius}}">
  <div class="mm-hr__badge"><span class="mm-hr__dot"></span><span class="mm-hr__state">—</span></div>
  <p class="mm-hr__detail"></p>
  <a class="mm-hr__btn" href="{{btnLink}}">{{btnText}}</a>
  <div class="mm-hr__data" hidden data-open="{{openLabel}}" data-close="{{closeLabel}}">
  {{#each hours}}
    {{!-- @name day @type outlined-textfield @default "월" @label "요일(월·화·수·목·금·토·일 중)" --}}
    {{!-- @name open @type time-picker @label "오픈 시간" --}}
    {{!-- @name close @type time-picker @label "마감 시간" --}}
    {{!-- @name off @type switch @default false @label "휴무" --}}
    <span data-day="{{day}}" data-o="{{open}}" data-c="{{close}}" data-off="{{off}}"></span>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-hr{max-width:380px;margin:0 auto;padding:clamp(20px,4vw,28px);text-align:center;border:1px solid #D4D0CB;background:#FAF9F7;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-hr__badge{display:inline-flex;align-items:center;gap:8px;padding:8px 16px;border-radius:999px;background:#F5F3F0;font-size:14px;font-weight:700;}
.mm-hr__dot{width:8px;height:8px;border-radius:50%;background:#B8B4AF;}
.mm-hr.is-open .mm-hr__dot{background:#1A1A1A;box-shadow:0 0 0 4px rgba(26,26,26,.12);}
.mm-hr.is-open .mm-hr__badge{background:#1A1A1A;color:#FAF9F7;}
.mm-hr__detail{margin:12px 0 0;font-size:13px;color:#8A8580;min-height:18px;}
.mm-hr__btn{display:inline-flex;align-items:center;justify-content:center;margin-top:16px;padding:13px 28px;border-radius:8px;background:#1A1A1A;color:#FAF9F7;font-weight:700;font-size:15px;text-decoration:none;transition:opacity .2s;}
.mm-hr__btn:empty,.mm-hr__btn[href=""]{display:none;}
@media (hover:hover){.mm-hr__btn:hover{opacity:.88;}}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-hr{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  var DAYS=['일','월','화','수','목','금','토'];
  function toMin(s){if(!s)return null;var pm=/오후|PM/i.test(s),am=/오전|AM/i.test(s);var m=String(s).match(/(\d{1,2})\D+(\d{2})/);if(!m)return null;var h=+m[1],mi=+m[2];if(pm&&h<12)h+=12;if(am&&h===12)h=0;return h*60+mi;}
  function initOne(root){
    var data=root.querySelector('.mm-hr__data');
    var stateEl=root.querySelector('.mm-hr__state');
    var detailEl=root.querySelector('.mm-hr__detail');
    if(!data)return;
    var openLabel=data.getAttribute('data-open')||'영업 중';
    var closeLabel=data.getAttribute('data-close')||'영업 종료';
    var rows=data.querySelectorAll('span');
    var now=new Date(),today=DAYS[now.getDay()],curMin=now.getHours()*60+now.getMinutes();
    var todayRow=null;
    for(var i=0;i<rows.length;i++){var d=rows[i].getAttribute('data-day')||'';if(d.indexOf(today)>=0){todayRow=rows[i];break;}}
    var isOpen=false,detail='';
    if(todayRow){
      var off=todayRow.getAttribute('data-off')==='true';
      var o=toMin(todayRow.getAttribute('data-o')),c=toMin(todayRow.getAttribute('data-c'));
      if(off){detail='오늘 휴무';}
      else if(o!=null&&c!=null){
        isOpen=curMin>=o&&curMin<c;
        function hm(x){var h=Math.floor(x/60),m=x%60;return (h<10?'0'+h:h)+':'+(m<10?'0'+m:m);}
        detail=hm(o)+' ~ '+hm(c);
      }
    }else{detail='오늘 영업정보 없음';}
    root.classList.toggle('is-open',isOpen);
    if(stateEl)stateEl.textContent=isOpen?openLabel:closeLabel;
    if(detailEl)detailEl.textContent=detail;
  }
  function init(){var l=document.querySelectorAll('.mm-hr');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
```

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

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-dx{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
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
  <div class="mm-360__stage" data-ar="{{ratio}}">
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

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-360{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
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

/* 동적 스타일(비율/배경/좌표): 인라인 style {{}}는 아임웹 저장거부 → data-* 속성을 JS로 적용 */
(function(){function ap(){
var A=document.querySelectorAll("[data-ar]");for(var i=0;i<A.length;i++){var v=(A[i].getAttribute("data-ar")||"").trim();if(v){if(A[i].classList.contains("mm-cat"))A[i].style.setProperty("--mm-ratio",v);else A[i].style.aspectRatio=v;}}
var B=document.querySelectorAll("[data-bg]");for(var i=0;i<B.length;i++){var v=(B[i].getAttribute("data-bg")||"").trim();if(v)B[i].style.backgroundImage="url('"+v+"')";}
var C=document.querySelectorAll("[data-x]");for(var i=0;i<C.length;i++){var x=(C[i].getAttribute("data-x")||"").trim(),y=(C[i].getAttribute("data-y")||"").trim();if(x)C[i].style.left=x+"%";if(y)C[i].style.top=y+"%";}
}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",ap);else ap();})();
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
  <div class="mm-opt__stage" data-ar="{{ratio}}"><img class="mm-opt__main" src="" alt=""></div>
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

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-opt{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
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

/* 동적 스타일(비율/배경/좌표): 인라인 style {{}}는 아임웹 저장거부 → data-* 속성을 JS로 적용 */
(function(){function ap(){
var A=document.querySelectorAll("[data-ar]");for(var i=0;i<A.length;i++){var v=(A[i].getAttribute("data-ar")||"").trim();if(v){if(A[i].classList.contains("mm-cat"))A[i].style.setProperty("--mm-ratio",v);else A[i].style.aspectRatio=v;}}
var B=document.querySelectorAll("[data-bg]");for(var i=0;i<B.length;i++){var v=(B[i].getAttribute("data-bg")||"").trim();if(v)B[i].style.backgroundImage="url('"+v+"')";}
var C=document.querySelectorAll("[data-x]");for(var i=0;i<C.length;i++){var x=(C[i].getAttribute("data-x")||"").trim(),y=(C[i].getAttribute("data-y")||"").trim();if(x)C[i].style.left=x+"%";if(y)C[i].style.top=y+"%";}
}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",ap);else ap();})();
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
{{!-- @name height @type outlined-textfield @default "" @label "높이 — 비우면 이미지 비율에 맞춰 자동(PC/모바일 각 이미지 기준). 값 넣으면 고정+크롭: 예 440px · 60vh" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name kicker @type outlined-textfield @default "CUT THE FAKE, KEEP THE REAL" @label "키커(영문)" --}}
{{!-- @name headline @type text-editor @default "<p>좋은 미용가위, 그 기준을 정의하다</p>" @label "헤드라인" --}}
{{!-- @name sub @type text-editor @default "<p></p>" @label "서브 문구(선택)" --}}
{{!-- @name btnText @type outlined-textfield @default "" @label "버튼 문구(선택)" --}}
{{!-- @name btnLink @type outlined-textfield @default "" @label "버튼 링크" --}}
<div class="mm-cut" data-radius="{{radius}}" data-theme="{{theme}}" data-height="{{height}}">
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
.mm-cut{box-sizing:border-box;position:relative;overflow:hidden;min-height:clamp(220px,42vw,440px);font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;background:#1A1A1A;--cut-s:clamp(.45,calc(var(--cut-h,320)/440),1);}
.mm-cut[data-theme="라이트"]{background:#FAF9F7;}
/* 기본=이미지 비율 자동: 흐름배치 → 컨테이너 높이=이미지 비율(PC=PC이미지 / 모바일=모바일이미지). 이미지 없으면 min-height 바닥 */
.mm-cut__bg,.mm-cut__bgm{display:block;width:100%;height:auto;opacity:.5;}
.mm-cut__bgm{display:none;}
@media (max-width:768px){.mm-cut[data-hasm="1"] .mm-cut__bg{display:none;}.mm-cut[data-hasm="1"] .mm-cut__bgm{display:block;}}
.mm-cut[data-theme="라이트"] .mm-cut__bg,.mm-cut[data-theme="라이트"] .mm-cut__bgm{opacity:.85;}
/* 고정 높이 모드(높이값 입력 시 JS가 data-fixed=1): 이미지 절대배치 cover-크롭, 높이는 min-height가 결정(텍스트가 안 밀어서 정확) */
.mm-cut[data-fixed="1"] .mm-cut__bg,.mm-cut[data-fixed="1"] .mm-cut__bgm{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;}
/* 카피는 이미지 위 오버레이(절대배치) → 이미지가 높이를 결정, 텍스트는 그 위 중앙. 폰트·간격은 배너 높이 비례(--cut-s)로 스케일 → 낮은 배너서 눌리지 않음 */
.mm-cut__inner{position:absolute;inset:0;z-index:1;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:calc(clamp(20px,5vw,52px) * var(--cut-s));}
.mm-cut__inner>*{max-width:760px;}
.mm-cut__kicker{display:block;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:calc(clamp(10px,2.6vw,12px) * var(--cut-s));font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#B8B4AF;opacity:0;transition:opacity .6s ease .1s;}
.mm-cut[data-theme="라이트"] .mm-cut__kicker{color:#8A8580;}
.mm-cut__blade{display:block;width:0;height:1px;margin:calc(16px * var(--cut-s)) auto;background:#FAF9F7;transition:width .7s cubic-bezier(.4,0,.2,1) .15s;}
.mm-cut[data-theme="라이트"] .mm-cut__blade{background:#1A1A1A;}
.mm-cut__headline{margin:0;font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:calc(clamp(24px,6.5vw,44px) * var(--cut-s));font-weight:900;line-height:1.2;letter-spacing:-.02em;color:#FAF9F7;opacity:0;transform:translateY(20px);transition:opacity .7s cubic-bezier(.4,0,.2,1) .35s,transform .7s cubic-bezier(.4,0,.2,1) .35s;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-cut[data-theme="라이트"] .mm-cut__headline{color:#1A1A1A;}
.mm-cut__sub{margin:calc(14px * var(--cut-s)) 0 0;font-size:calc(clamp(14px,3.8vw,17px) * var(--cut-s));line-height:1.5;color:#D4D0CB;opacity:0;transition:opacity .7s ease .55s;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-cut[data-theme="라이트"] .mm-cut__sub{color:#4A4A4A;}
.mm-cut__sub:empty{display:none;}
.mm-cut__btn{display:inline-flex;margin-top:calc(22px * var(--cut-s));padding:calc(13px * var(--cut-s)) calc(32px * var(--cut-s));border-radius:8px;background:#FAF9F7;color:#1A1A1A;font-weight:700;font-size:calc(15px * var(--cut-s));text-decoration:none;opacity:0;transition:opacity .7s ease .7s,transform .2s;}
.mm-cut[data-theme="라이트"] .mm-cut__btn{background:#1A1A1A;color:#FAF9F7;}
.mm-cut__btn:empty{display:none;}
.mm-cut__btn:active{transform:scale(.97);}
.mm-cut.is-in .mm-cut__kicker,.mm-cut.is-in .mm-cut__headline,.mm-cut.is-in .mm-cut__sub,.mm-cut.is-in .mm-cut__btn{opacity:1;transform:none;}
.mm-cut.is-in .mm-cut__blade{width:calc(clamp(40px,12vw,90px) * var(--cut-s));}
/* 여러 줄(text-editor) 입력 시 문단 줄간격 통일 */
.mm-cut__headline p,.mm-cut__sub p{margin:0 0 .15em;}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var m=root.querySelector('.mm-cut__bgm');
    if(m && String(m.getAttribute('src')||'').trim()) root.setAttribute('data-hasm','1');
    /* 높이: 비우면 이미지 비율 자동(CSS). 값 입력 시 고정 높이+크롭(data-fixed) → 텍스트가 안 밀어 정확 */
    var h=root.getAttribute('data-height');
    if(h && h.trim()){ var hv=h.trim(); if(String(parseFloat(hv))===hv) hv+='px'; root.style.minHeight=hv; root.setAttribute('data-fixed','1'); }
    else { root.removeAttribute('data-fixed'); root.style.minHeight=''; }
    /* 배너 실제 높이 → --cut-h 주입 → CSS가 폰트·간격을 높이 비례로 스케일(낮은 배너서 안 눌림). 이미지로드·리사이즈·높이변경 자동 반영 */
    function applyScale(){ var hh=Math.round(root.getBoundingClientRect().height); if(hh>0) root.style.setProperty('--cut-h', hh); }
    applyScale();
    if('ResizeObserver' in window){ new ResizeObserver(applyScale).observe(root); }
    else { window.addEventListener('resize', applyScale); }
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('is-in');io.unobserve(e.target);}});},{threshold:.3});
      io.observe(root);
    }else{root.classList.add('is-in');}
  }
  function init(){var l=document.querySelectorAll('.mm-cut');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
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

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-gate{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
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

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-daily{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
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
{{!-- @name height @type outlined-textfield @default "" @label "높이 — 비우면 이미지 비율에 맞춰 자동(권장). 값 넣으면 고정+크롭: 예 480px · 60vh" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name maxw @type outlined-textfield @default "" @label "최대 가로폭(PC) — 숫자 자유 입력: 예 1280 (비우거나 full=꽉 채움). 영역 확장해도 이 값에서 멈춤 · 모바일 자동 꽉 채움" --}}
{{!-- @name focus @type outlined-textfield @default "중앙" @label "사진 초점 — 입력: 중앙·좌·우·상·하" --}}
{{!-- @name overlay @type color-picker @default "#1A1A1A80" @label "어둡게 (검정의 투명도 슬라이더를 드래그 · 맨뒤 2자리=어둡기)" --}}
{{!-- @name align @type outlined-textfield @default "가운데" @label "정렬 (가운데/왼쪽 · 유형을 옵션버튼/스위치로 바꿔도 됨)" --}}
{{!-- @name kicker @type outlined-textfield @default "" @label "키커(영문, 선택)" --}}
{{!-- @name headline @type text-editor @default "<p>헤드라인을 입력하세요</p>" @label "헤드라인" --}}
{{!-- @name sub @type text-editor @default "<p></p>" @label "서브 문구(선택)" --}}
{{!-- @name btnText @type outlined-textfield @default "" @label "버튼 문구(선택)" --}}
{{!-- @name btnLink @type outlined-textfield @default "" @label "버튼 링크" --}}
<div class="mm-cine" data-radius="{{radius}}" data-align="{{align}}" data-overlay="{{overlay}}" data-focus="{{focus}}" data-height="{{height}}" data-maxw="{{maxw}}">
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
/* 가로 최대폭은 자유 숫자값(예 1280) → JS가 max-width 적용(margin auto로 중앙). 영역 확장해도 이 값에서 멈춤 */
.mm-cine{box-sizing:border-box;margin-left:auto;margin-right:auto;position:relative;overflow:hidden;background:#1A1A1A;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
/* 기본=이미지 비율 자동: 이미지가 흐름에 들어가 컨테이너 높이=이미지 비율(PC=PC이미지 / 모바일=모바일이미지). 높이 미입력 시 이 모드 */
.mm-cine__bg,.mm-cine__bgm{display:block;width:100%;height:auto;animation:mm-cine-zoom 16s ease-out both;}
.mm-cine__bgm{display:none;}
/* 모바일 전용 이미지가 있을 때(data-hasm=1)만 모바일에서 스왑 → 모바일 높이=모바일 이미지 비율 */
@media (max-width:768px){
  .mm-cine[data-hasm="1"] .mm-cine__bg{display:none;}
  .mm-cine[data-hasm="1"] .mm-cine__bgm{display:block;}
}
/* 고정 높이 모드(높이값 입력 시 JS가 data-fixed=1): 이미지 절대배치 cover-크롭, 높이는 min-height가 결정 */
.mm-cine[data-fixed="1"] .mm-cine__bg,.mm-cine[data-fixed="1"] .mm-cine__bgm{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;}
@keyframes mm-cine-zoom{from{transform:scale(1.0)}to{transform:scale(1.09)}}
.mm-cine__veil{position:absolute;inset:0;z-index:1;background:rgba(26,26,26,.45);}
/* 카피는 이미지 위 오버레이(절대배치) → 이미지가 높이를 결정하고 텍스트는 그 위에 얹힘 */
.mm-cine__inner{position:absolute;inset:0;z-index:2;display:flex;flex-direction:column;justify-content:center;padding:clamp(28px,6vw,56px);}
.mm-cine__inner>*{max-width:720px;}
.mm-cine[data-align="가운데"] .mm-cine__inner{align-items:center;text-align:center;}
.mm-cine[data-align="왼쪽"] .mm-cine__inner{align-items:flex-start;text-align:left;}
.mm-cine__kicker{display:block;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(10px,2.6vw,12px);font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#D4D0CB;margin-bottom:12px;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-cine__kicker:empty{display:none;}
.mm-cine__headline{margin:0;font-family:'Outfit','Plus Jakarta Sans','Pretendard','Noto Sans KR',sans-serif;font-size:clamp(24px,6vw,44px);font-weight:900;line-height:1.12;letter-spacing:-.02em;color:#FAF9F7;text-shadow:0 2px 24px rgba(0,0,0,.35);white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-cine__headline p{margin:0 0 .15em;}
.mm-cine__sub{margin:14px 0 0;font-size:clamp(14px,3.8vw,17px);line-height:1.45;color:#EDEBE8;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-cine__sub p{margin:0 0 .15em;}
.mm-cine__sub:empty{display:none;}
.mm-cine__btn{display:inline-flex;margin-top:24px;padding:14px 34px;border-radius:8px;background:#FAF9F7;color:#1A1A1A;font-weight:700;font-size:15px;text-decoration:none;transition:opacity .2s,transform .2s;}
.mm-cine__btn:empty{display:none;}
.mm-cine__btn:active{transform:scale(.97);}
@media (hover:hover){.mm-cine__btn:hover{opacity:.9;}}
@media (prefers-reduced-motion:reduce){.mm-cine__bg,.mm-cine__bgm{animation:none;}}
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
  /* 가로 최대폭: 자유 숫자값(1280 등) 적용. 비움/full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
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
    /* 높이: 비우면 이미지 비율 자동(CSS가 담당). 값 입력 시 고정 높이+크롭 모드(data-fixed) */
    var h=root.getAttribute('data-height');
    if(h && h.trim()){ var hv=h.trim(); if(String(parseFloat(hv))===hv) hv+='px'; root.style.minHeight=hv; root.setAttribute('data-fixed','1'); }
    else { root.removeAttribute('data-fixed'); root.style.minHeight=''; }
    var rd=root.getAttribute('data-radius'); if(rd!==null){ var rv=rd.trim(); if(rv){ if(String(parseFloat(rv))===rv) rv+='px'; root.style.borderRadius=rv; } }
    /* 가로 최대폭 적용 + 패널값 바뀌면(data-maxw 속성변경) 즉시 재적용 → 편집기 실시간 반영 */
    applyMaxw(root);
    if('MutationObserver' in window){ new MutationObserver(function(){applyMaxw(root);}).observe(root,{attributes:true,attributeFilter:['data-maxw']}); }
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
<div class="mm-split" data-radius="{{radius}}" data-img="{{imgPos}}" data-theme="{{theme}}" data-height="{{height}}">
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
.mm-split{display:grid;grid-template-columns:1fr 1fr;align-items:stretch;overflow:hidden;border:1px solid #D4D0CB;background:#FFFFFF;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
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

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-split{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
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

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
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
<div class="mm-notice" data-radius="{{radius}}" data-theme="{{theme}}" data-flow="{{flow}}">
  <div class="mm-notice__in">
    <span class="mm-notice__icon">{{icon}}</span>
    <span class="mm-notice__text">{{text}}</span>
    <a class="mm-notice__link" href="{{link}}">{{linkText}}</a>
  </div>
</div>
```
### CSS 탭
```css
.mm-notice{width:100%;background:#1A1A1A;color:#FAF9F7;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;overflow:hidden;}
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

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-notice{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
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

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
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
<div class="mm-quote" data-radius="{{radius}}" data-theme="{{theme}}" data-align="{{align}}">
  <span class="mm-quote__mark" aria-hidden="true">“</span>
  <blockquote class="mm-quote__text">{{quote}}</blockquote>
  <cite class="mm-quote__author">{{author}}</cite>
</div>
```
### CSS 탭
```css
.mm-quote{position:relative;max-width:860px;margin:0 auto;padding:clamp(36px,7vw,72px) clamp(20px,5vw,48px);background:#FAF9F7;color:#1A1A1A;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
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

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-quote{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
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

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
```

---

## D5. 듀얼 선택 배너
`폴더: d05_dual_choice`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "두 갈래 이미지 타일 선택 배너." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 듀얼 선택 배너
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 두 갈래 길(제품/복원수리 등)을 큰 이미지 타일 2개로 (디자인+내비)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name tiles @type item @label "타일(2개 권장)" --}}
<div class="mm-dual">
{{#each tiles}}
  {{!-- @name image @type image @label "타일 이미지 (권장 800×1000px)" --}}
  {{!-- @name label @type outlined-textfield @default "선택" @label "큰 라벨" --}}
  {{!-- @name desc @type outlined-textfield @default "" @label "작은 설명(선택)" --}}
  {{!-- @name link @type outlined-textfield @default "" @label "링크" --}}
  <a class="mm-dual__tile" href="{{link}}">
    <img class="mm-dual__img" src="{{image}}" alt="">
    <span class="mm-dual__veil" aria-hidden="true"></span>
    <span class="mm-dual__cap">
      <span class="mm-dual__label">{{label}}</span>
      <span class="mm-dual__desc">{{desc}}</span>
      <span class="mm-dual__arrow" aria-hidden="true">→</span>
    </span>
  </a>
{{/each}}
</div>
```
### CSS 탭
```css
.mm-dual{max-width:920px;margin:0 auto;display:grid;grid-template-columns:1fr 1fr;gap:14px;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-dual__tile{position:relative;display:block;overflow:hidden;border-radius:16px;background:#1A1A1A;aspect-ratio:3/4;text-decoration:none;}
.mm-dual__img{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;transition:transform .5s cubic-bezier(.4,0,.2,1);}
.mm-dual__veil{position:absolute;inset:0;background:linear-gradient(rgba(26,26,26,0) 35%,rgba(26,26,26,.82));}
.mm-dual__cap{position:absolute;left:0;right:0;bottom:0;padding:clamp(18px,3vw,26px);color:#FAF9F7;display:flex;flex-direction:column;gap:4px;}
.mm-dual__label{font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:clamp(18px,4vw,26px);font-weight:800;letter-spacing:-.02em;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-dual__desc{font-size:clamp(12px,3vw,14px);color:#D4D0CB;font-weight:600;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-dual__desc:empty{display:none;}
.mm-dual__arrow{margin-top:8px;font-size:18px;transition:transform .3s cubic-bezier(.4,0,.2,1);}
@media (hover:hover){.mm-dual__tile:hover .mm-dual__img{transform:scale(1.06);}.mm-dual__tile:hover .mm-dual__arrow{transform:translateX(5px);}}
@media (max-width:520px){.mm-dual{grid-template-columns:1fr;}.mm-dual__tile{aspect-ratio:16/10;}}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-dual{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
/* 듀얼 선택 배너 — 순수 HTML/CSS 동작. JS 불필요(빈 안전망). */
(function(){
  function init(){var l=document.querySelectorAll('.mm-dual');if(!l.length){return setTimeout(init,50);}}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## N1. 가위 길이 시뮬레이터
`폴더: n01_size_ruler`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "슬라이더로 인치 선택, cm 환산·길이 바." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 가위 길이 시뮬레이터
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 슬라이더로 인치 선택 → 길이 비교 바 + cm 환산 (사이즈 이해)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name min @type outlined-textfield @default "4.5" @label "최소 인치" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name max @type outlined-textfield @default "7.0" @label "최대 인치" --}}
{{!-- @name def @type outlined-textfield @default "5.5" @label "기본 인치" --}}
{{!-- @name note @type outlined-textfield @default "* 화면 비례 참고용이며 실제 길이와 다를 수 있어요" @label "안내(선택)" --}}
<div class="mm-rul" data-radius="{{radius}}" data-min="{{min}}" data-max="{{max}}" data-def="{{def}}">
  <div class="mm-rul__readout"><strong class="mm-rul__inch">—</strong><span class="mm-rul__cm">—</span></div>
  <div class="mm-rul__track"><div class="mm-rul__bar"></div></div>
  <input type="range" class="mm-rul__range" aria-label="가위 길이(인치)">
  <p class="mm-rul__note">{{note}}</p>
</div>
```
### CSS 탭
```css
.mm-rul{max-width:520px;margin:0 auto;padding:clamp(22px,5vw,32px);border:1px solid #D4D0CB;background:#FAF9F7;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;text-align:center;}
.mm-rul__readout{display:flex;align-items:baseline;justify-content:center;gap:10px;margin-bottom:18px;}
.mm-rul__inch{font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(34px,10vw,52px);font-weight:900;letter-spacing:-.03em;line-height:1;}
.mm-rul__cm{font-size:clamp(14px,4vw,17px);font-weight:600;color:#8A8580;}
.mm-rul__track{height:14px;background:#EDEBE8;border-radius:999px;overflow:hidden;margin-bottom:18px;}
.mm-rul__bar{height:100%;width:50%;background:#1A1A1A;border-radius:999px;transition:width .15s cubic-bezier(.4,0,.2,1);}
.mm-rul__range{width:100%;-webkit-appearance:none;appearance:none;height:4px;border-radius:999px;background:#D4D0CB;outline:none;}
.mm-rul__range::-webkit-slider-thumb{-webkit-appearance:none;width:26px;height:26px;border-radius:50%;background:#1A1A1A;cursor:pointer;border:3px solid #FAF9F7;box-shadow:0 2px 8px rgba(0,0,0,.25);}
.mm-rul__range::-moz-range-thumb{width:26px;height:26px;border-radius:50%;background:#1A1A1A;cursor:pointer;border:3px solid #FAF9F7;}
.mm-rul__note{margin:14px 0 0;font-size:12px;color:#B8B4AF;}
.mm-rul__note:empty{display:none;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-rul{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function num(v,d){var n=parseFloat(v);return isNaN(n)?d:n;}
  function initOne(root){
    var range=root.querySelector('.mm-rul__range');
    var bar=root.querySelector('.mm-rul__bar');
    var inchEl=root.querySelector('.mm-rul__inch');
    var cmEl=root.querySelector('.mm-rul__cm');
    if(!range)return;
    var min=num(root.getAttribute('data-min'),4.5),max=num(root.getAttribute('data-max'),7.0),def=num(root.getAttribute('data-def'),5.5);
    range.min=min;range.max=max;range.step=0.1;range.value=Math.min(max,Math.max(min,def));
    function upd(){
      var v=parseFloat(range.value);
      var p=max>min?(v-min)/(max-min):0;
      if(bar)bar.style.width=(20+p*80)+'%';
      if(inchEl)inchEl.textContent=v.toFixed(1)+'″';
      if(cmEl)cmEl.textContent='('+(v*2.54).toFixed(1)+'cm)';
    }
    range.addEventListener('input',upd);upd();
  }
  function init(){var l=document.querySelectorAll('.mm-rul');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
```

---

## N2. 마모루 vs 일반 비교표
`폴더: n02_versus_table`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "항목별 마모루/일반 비교(O·X는 자동 ✓·✕)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 마모루 vs 일반 비교표
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 항목별 마모루 / 일반 판매처 비교 (자체복원 등 차별점 강조)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name leftLabel @type outlined-textfield @default "MAMORU" @label "왼쪽(자사) 라벨" --}}
{{!-- @name rightLabel @type outlined-textfield @default "일반 판매처" @label "오른쪽 라벨" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name rows @type item @label "비교 항목" --}}
<div class="mm-vs" data-radius="{{radius}}">
  <div class="mm-vs__head">
    <span class="mm-vs__h mm-vs__h--label"></span>
    <span class="mm-vs__h mm-vs__h--us">{{leftLabel}}</span>
    <span class="mm-vs__h mm-vs__h--them">{{rightLabel}}</span>
  </div>
  {{#each rows}}
    {{!-- @name label @type outlined-textfield @default "항목" @label "항목명" --}}
    {{!-- @name us @type outlined-textfield @default "O" @label "자사(O/X 또는 텍스트)" --}}
    {{!-- @name them @type outlined-textfield @default "X" @label "일반(O/X 또는 텍스트)" --}}
    <div class="mm-vs__row">
      <span class="mm-vs__label">{{label}}</span>
      <span class="mm-vs__cell mm-vs__cell--us" data-v="{{us}}"></span>
      <span class="mm-vs__cell mm-vs__cell--them" data-v="{{them}}"></span>
    </div>
  {{/each}}
</div>
```
### CSS 탭
```css
.mm-vs{max-width:600px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;border:1px solid #D4D0CB;overflow:hidden;}
.mm-vs__head,.mm-vs__row{display:grid;grid-template-columns:1.4fr 1fr 1fr;align-items:center;}
.mm-vs__head{background:#1A1A1A;color:#FAF9F7;}
.mm-vs__h{padding:14px 10px;font-size:clamp(12px,3.2vw,14px);font-weight:700;text-align:center;}
.mm-vs__h--us{font-family:'Outfit','Plus Jakarta Sans',sans-serif;letter-spacing:.04em;}
.mm-vs__h--them{color:#B8B4AF;font-weight:600;}
.mm-vs__row{border-top:1px solid #EDEBE8;}
.mm-vs__row:nth-child(even){background:#FAF9F7;}
.mm-vs__label{padding:14px 14px;font-size:clamp(13px,3.6vw,15px);font-weight:600;}
.mm-vs__cell{padding:14px 10px;text-align:center;font-size:clamp(13px,3.6vw,15px);font-weight:700;}
.mm-vs__cell--us{color:#1A1A1A;}
.mm-vs__cell--them{color:#B8B4AF;}
.mm-vs__cell.is-o::after{content:'✓';font-size:18px;}
.mm-vs__cell.is-x::after{content:'✕';font-size:16px;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-vs{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function render(cell){
    var v=(cell.getAttribute('data-v')||'').trim();
    var low=v.toLowerCase();
    if(low==='o'||low==='yes'||low==='y'||v==='✓'||v==='○'){cell.classList.add('is-o');cell.textContent='';}
    else if(low==='x'||low==='no'||low==='n'||v==='✕'||v==='×'){cell.classList.add('is-x');cell.textContent='';}
    else{cell.textContent=v;}
  }
  function initOne(root){var cells=root.querySelectorAll('.mm-vs__cell');for(var i=0;i<cells.length;i++)render(cells[i]);}
  function init(){var l=document.querySelectorAll('.mm-vs');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
```

---

## N3. 가로 스크롤 핀 갤러리
`폴더: n03_horizontal_pin`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "세로 스크롤 시 가로로 흐르는 핀 갤러리(전체너비)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 가로 스크롤 핀 갤러리
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 세로 스크롤 시 화면이 고정된 채 콘텐츠가 가로로 흐름 (트렌디 쇼케이스)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name slides @type item @label "슬라이드" --}}
<div class="mm-hpin">
  <div class="mm-hpin__sticky">
    <div class="mm-hpin__track">
    {{#each slides}}
      {{!-- @name image @type image @label "이미지 (권장 1600×900px)" --}}
      {{!-- @name caption @type outlined-textfield @default "" @label "캡션(선택)" --}}
      <div class="mm-hpin__panel">
        <img class="mm-hpin__img" src="{{image}}" alt="">
        <p class="mm-hpin__cap">{{caption}}</p>
      </div>
    {{/each}}
    </div>
  </div>
</div>
```
### CSS 탭
```css
.mm-hpin{position:relative;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-hpin__sticky{position:sticky;top:0;height:100vh;overflow:hidden;display:flex;align-items:center;}
.mm-hpin__track{display:flex;gap:clamp(16px,3vw,32px);padding:0 clamp(16px,4vw,48px);will-change:transform;}
.mm-hpin__panel{flex:0 0 auto;width:min(78vw,560px);}
.mm-hpin__img{width:100%;aspect-ratio:4/3;object-fit:cover;border-radius:16px;background:#F5F3F0;display:block;}
.mm-hpin__cap{margin:14px 2px 0;font-size:clamp(14px,3.6vw,16px);font-weight:600;color:#4A4A4A;}
.mm-hpin__cap:empty{display:none;}
/* 핀 불가/모바일·축소모션: 네이티브 가로 스크롤로 폴백 */
.mm-hpin.is-native .mm-hpin__sticky{position:static;height:auto;overflow-x:auto;scroll-snap-type:x mandatory;-webkit-overflow-scrolling:touch;}
.mm-hpin.is-native .mm-hpin__track{transform:none!important;}
.mm-hpin.is-native .mm-hpin__panel{scroll-snap-align:center;}
.mm-hpin.is-native .mm-hpin__sticky::-webkit-scrollbar{display:none;}
@media (prefers-reduced-motion:reduce){.mm-hpin__sticky{position:static;height:auto;overflow-x:auto;}.mm-hpin__track{transform:none!important;}}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-hpin{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var sticky=root.querySelector('.mm-hpin__sticky');
    var track=root.querySelector('.mm-hpin__track');
    if(!sticky||!track)return;
    var reduce=window.matchMedia&&window.matchMedia('(prefers-reduced-motion:reduce)').matches;
    var touch=window.matchMedia&&window.matchMedia('(hover:none)').matches;
    function setup(){
      var dist=track.scrollWidth-window.innerWidth;
      if(reduce||touch||dist<=40){root.classList.add('is-native');root.style.height='';track.style.transform='';return;}
      root.classList.remove('is-native');
      root.style.height=(dist+window.innerHeight)+'px';
      onScroll();
    }
    function onScroll(){
      if(root.classList.contains('is-native'))return;
      var rect=root.getBoundingClientRect();
      var dist=track.scrollWidth-window.innerWidth;
      var p=Math.min(Math.max(-rect.top/(root.offsetHeight-window.innerHeight),0),1);
      track.style.transform='translateX('+(-p*dist)+'px)';
    }
    window.addEventListener('scroll',onScroll,{passive:true});
    window.addEventListener('resize',setup);
    setup();
    // 이미지 로드 후 폭 재계산
    var imgs=root.querySelectorAll('img');for(var i=0;i<imgs.length;i++)imgs[i].addEventListener('load',setup);
  }
  function init(){var l=document.querySelectorAll('.mm-hpin');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## N4. 텍스트 마스크 헤드라인
`폴더: n04_text_mask`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "큰 글자 안에 이미지가 비치는 헤드라인." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 텍스트 마스크 헤드라인
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 큰 글자 안으로 이미지가 비쳐 보이는 타이포 헤드라인 (프리미엄)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name image @type image @label "글자 속 이미지 — 권장 1200×600px (대비 강한 이미지)" --}}
{{!-- @name headline @type text-editor @default "<p>MAMORU</p>" @label "헤드라인(짧고 굵게)" --}}
{{!-- @name sub @type text-editor @default "<p></p>" @label "아래 문구(선택)" --}}
{{!-- @name align @type outlined-textfield @default "가운데" @label "정렬 — 입력: 가운데 · 왼쪽" --}}
<div class="mm-mask" data-align="{{align}}">
  <div class="mm-mask__h" data-bg="{{image}}">{{headline}}</div>
  <div class="mm-mask__sub">{{sub}}</div>
</div>
```
### CSS 탭
```css
.mm-mask{max-width:900px;margin:0 auto;padding:clamp(24px,5vw,40px) 16px;text-align:center;font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;}
.mm-mask[data-align="왼쪽"]{text-align:left;}
.mm-mask__h{margin:0;font-size:clamp(48px,18vw,140px);font-weight:900;letter-spacing:-.03em;line-height:.95;
  background-size:cover;background-position:center;background-repeat:no-repeat;
  -webkit-background-clip:text;background-clip:text;color:transparent;-webkit-text-fill-color:transparent;
  white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
@supports not ((-webkit-background-clip:text) or (background-clip:text)){
  .mm-mask__h{color:#1A1A1A;-webkit-text-fill-color:#1A1A1A;background:none;}
}
.mm-mask__sub{margin:clamp(10px,2vw,18px) 4px 0;font-family:'Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:clamp(14px,3.8vw,17px);font-weight:600;color:#8A8580;letter-spacing:.01em;}
.mm-mask__sub:empty{display:none;}

/* 여러 줄(text-editor) 입력 시 문단 줄간격 통일 */
.mm-mask__h p,.mm-mask__sub p{margin:0 0 .15em;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-mask{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
/* 텍스트 마스크 — 정렬값 관대 인식 + 순수 CSS(background-clip:text). 진입 reveal. */
(function(){
  function isLeft(al){al=String(al||'').trim();var low=al.toLowerCase();return al==='true'||al.indexOf('왼')>=0||al.indexOf('좌')>=0||low.indexOf('left')>=0||low.indexOf('start')>=0;}
  function init(){
    var l=document.querySelectorAll('.mm-mask');
    if(!l.length){return setTimeout(init,50);}
    for(var i=0;i<l.length;i++) l[i].setAttribute('data-align', isLeft(l[i].getAttribute('data-align'))?'왼쪽':'가운데');
    if(!('IntersectionObserver' in window))return;
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.style.opacity='1';e.target.style.transform='none';io.unobserve(e.target);}});},{threshold:.2});
    for(var j=0;j<l.length;j++){l[j].style.transition='opacity .7s cubic-bezier(.4,0,.2,1),transform .7s cubic-bezier(.4,0,.2,1)';l[j].style.opacity='0';l[j].style.transform='translateY(20px)';io.observe(l[j]);}
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 동적 스타일(비율/배경/좌표): 인라인 style {{}}는 아임웹 저장거부 → data-* 속성을 JS로 적용 */
(function(){function ap(){
var A=document.querySelectorAll("[data-ar]");for(var i=0;i<A.length;i++){var v=(A[i].getAttribute("data-ar")||"").trim();if(v){if(A[i].classList.contains("mm-cat"))A[i].style.setProperty("--mm-ratio",v);else A[i].style.aspectRatio=v;}}
var B=document.querySelectorAll("[data-bg]");for(var i=0;i<B.length;i++){var v=(B[i].getAttribute("data-bg")||"").trim();if(v)B[i].style.backgroundImage="url('"+v+"')";}
var C=document.querySelectorAll("[data-x]");for(var i=0;i<C.length;i++){var x=(C[i].getAttribute("data-x")||"").trim(),y=(C[i].getAttribute("data-y")||"").trim();if(x)C[i].style.left=x+"%";if(y)C[i].style.top=y+"%";}
}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",ap);else ap();})();
```

---

## N5. 한정 수량 게이지
`폴더: n05_stock_gauge`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "남은 수량 게이지(패널 입력값, 실시간 아님)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 한정 수량 게이지
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 남은 수량을 게이지 바로 (희소성·긴박감, 전환↑)
  🚫 fetch·iframe 0 (※실시간 재고 아님 — 패널 입력값)
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "한정 수량" @label "제목" --}}
{{!-- @name total @type outlined-textfield @default "100" @label "전체 수량" --}}
{{!-- @name remaining @type outlined-textfield @default "23" @label "남은 수량" --}}
{{!-- @name unit @type outlined-textfield @default "개" @label "단위" --}}
{{!-- @name radius @type outlined-textfield @default "14px" @label "모서리 둥글기 — 예: 14px · 0px이면 각지게" --}}
<div class="mm-stock" data-radius="{{radius}}" data-total="{{total}}" data-remaining="{{remaining}}" data-unit="{{unit}}">
  <div class="mm-stock__top"><span class="mm-stock__title">{{title}}</span><span class="mm-stock__count">—</span></div>
  <div class="mm-stock__track"><div class="mm-stock__bar"></div></div>
</div>
```
### CSS 탭
```css
.mm-stock{max-width:440px;margin:0 auto;padding:clamp(18px,4vw,24px);border:1px solid #D4D0CB;background:#FAF9F7;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-stock__top{display:flex;align-items:baseline;justify-content:space-between;margin-bottom:12px;}
.mm-stock__title{font-size:14px;font-weight:700;}
.mm-stock__count{font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(15px,4vw,18px);font-weight:800;letter-spacing:-.01em;}
.mm-stock__count b{font-size:1.25em;}
.mm-stock__track{height:12px;background:#EDEBE8;border-radius:999px;overflow:hidden;}
.mm-stock__bar{height:100%;width:0;background:#1A1A1A;border-radius:999px;transition:width 1s cubic-bezier(.4,0,.2,1);}
.mm-stock.is-low .mm-stock__bar{background:#2D2D2D;}
.mm-stock.is-low .mm-stock__count{color:#1A1A1A;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-stock{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function num(v,d){var n=parseFloat(String(v).replace(/[^0-9.]/g,''));return isNaN(n)?d:n;}
  function initOne(root){
    var total=num(root.getAttribute('data-total'),100);
    var rem=num(root.getAttribute('data-remaining'),0);
    var unit=root.getAttribute('data-unit')||'';
    var bar=root.querySelector('.mm-stock__bar');
    var count=root.querySelector('.mm-stock__count');
    rem=Math.max(0,Math.min(rem,total));
    var pct=total>0?rem/total*100:0;
    if(count)count.innerHTML='<b>'+rem+'</b>'+unit+' 남음';
    root.classList.toggle('is-low',pct<=30);
    function go(){if(bar)bar.style.width=pct+'%';}
    if('IntersectionObserver' in window){var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){go();io.disconnect();}});},{threshold:.4});io.observe(root);}else{go();}
  }
  function init(){var l=document.querySelectorAll('.mm-stock');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
```

---

## N6. 이용 안내 가로 스텝
`폴더: n06_steps`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "이용 절차를 번호 가로 스텝으로(모바일 세로)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 이용 안내 가로 스텝
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 상담→진단→맞춤→배송 등 절차를 번호·연결선 가로 스텝으로 (명확성)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "" @label "제목(선택)" --}}
{{!-- @name steps @type item @label "단계" --}}
<div class="mm-step">
  <p class="mm-step__title">{{title}}</p>
  <ol class="mm-step__list">
  {{#each steps}}
    {{!-- @name name @type outlined-textfield @default "단계" @label "단계명" --}}
    {{!-- @name desc @type outlined-textfield @default "" @label "설명(선택)" --}}
    <li class="mm-step__item">
      <span class="mm-step__num" aria-hidden="true"></span>
      <span class="mm-step__name">{{name}}</span>
      <span class="mm-step__desc">{{desc}}</span>
    </li>
  {{/each}}
  </ol>
</div>
```
### CSS 탭
```css
.mm-step{max-width:880px;margin:0 auto;padding:clamp(20px,4vw,32px) 8px;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-step__title{margin:0 0 24px;text-align:center;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(20px,5vw,28px);font-weight:800;letter-spacing:-.02em;}
.mm-step__title:empty{display:none;}
.mm-step__list{counter-reset:st;list-style:none;margin:0;padding:0;display:flex;gap:0;}
.mm-step__item{counter-increment:st;flex:1;position:relative;display:flex;flex-direction:column;align-items:center;text-align:center;padding:0 8px;}
.mm-step__item::before{content:'';position:absolute;top:22px;left:-50%;width:100%;height:1px;background:#D4D0CB;z-index:0;}
.mm-step__item:first-child::before{display:none;}
.mm-step__num{position:relative;z-index:1;width:44px;height:44px;border-radius:50%;background:#1A1A1A;color:#FAF9F7;display:flex;align-items:center;justify-content:center;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:17px;font-weight:800;margin-bottom:12px;}
.mm-step__num::before{content:counter(st);}
.mm-step__name{font-size:clamp(13px,3.4vw,15px);font-weight:700;letter-spacing:-.01em;}
.mm-step__desc{margin-top:4px;font-size:clamp(11px,3vw,13px);color:#8A8580;line-height:1.4;}
.mm-step__desc:empty{display:none;}
@media (max-width:560px){
  .mm-step__list{flex-direction:column;gap:0;}
  .mm-step__item{flex-direction:row;align-items:flex-start;text-align:left;gap:14px;padding:0 0 24px;}
  .mm-step__item::before{top:44px;left:21px;width:1px;height:100%;}
  .mm-step__num{margin-bottom:0;flex:0 0 auto;}
  .mm-step__name,.mm-step__desc{align-self:center;}
  .mm-step__item{display:grid;grid-template-columns:44px 1fr;row-gap:2px;}
  .mm-step__name{grid-column:2;}.mm-step__desc{grid-column:2;}
}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-step{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
/* 이용 안내 가로 스텝 — CSS 카운터로 번호 자동. 진입 reveal(선택). */
(function(){
  function init(){
    var l=document.querySelectorAll('.mm-step');
    if(!l.length){return setTimeout(init,50);}
    if(!('IntersectionObserver' in window))return;
    for(var w=0;w<l.length;w++){
      var items=l[w].querySelectorAll('.mm-step__item');
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.style.opacity='1';e.target.style.transform='none';io.unobserve(e.target);}});},{threshold:.25});
      for(var i=0;i<items.length;i++){items[i].style.transition='opacity .5s cubic-bezier(.4,0,.2,1) '+(i*0.08)+'s, transform .5s cubic-bezier(.4,0,.2,1) '+(i*0.08)+'s';items[i].style.opacity='0';items[i].style.transform='translateY(14px)';io.observe(items[i]);}
    }
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## N7. 에디토리얼 섹션 헤더
`폴더: n07_section_header`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "거대 인덱스+키커+제목 섹션 구분자." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 에디토리얼 섹션 헤더
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 거대 인덱스 + 키커 + 제목 + hairline (B-08 에디토리얼 섹션 구분)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name index @type outlined-textfield @default "01" @label "인덱스(예: 01)" --}}
{{!-- @name kicker @type outlined-textfield @default "SECTION" @label "키커(영문)" --}}
{{!-- @name title @type text-editor @default "<p>섹션 제목</p>" @label "제목" --}}
{{!-- @name align @type outlined-textfield @default "왼쪽" @label "정렬 — 입력: 왼쪽 · 가운데" --}}
<div class="mm-sh" data-align="{{align}}">
  <span class="mm-sh__index">{{index}}</span>
  <div class="mm-sh__text">
    <span class="mm-sh__kicker">{{kicker}}</span>
    <div class="mm-sh__title" role="heading" aria-level="2">{{title}}</div>
  </div>
  <span class="mm-sh__line" aria-hidden="true"></span>
</div>
```
### CSS 탭
```css
.mm-sh{max-width:900px;margin:0 auto;padding:clamp(20px,4vw,32px) 8px;display:flex;align-items:baseline;gap:clamp(14px,3vw,28px);font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-sh[data-align="가운데"]{flex-direction:column;align-items:center;text-align:center;gap:8px;}
.mm-sh__index{font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(40px,11vw,76px);font-weight:900;letter-spacing:-.05em;line-height:.9;color:#1A1A1A;flex:0 0 auto;}
.mm-sh__text{flex:0 1 auto;}
.mm-sh__kicker{display:block;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(10px,2.6vw,12px);font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#8A8580;margin-bottom:6px;}
.mm-sh__title{margin:0;font-size:clamp(20px,5.5vw,32px);font-weight:800;letter-spacing:-.02em;line-height:1.2;}
.mm-sh__line{flex:1;height:1px;background:#D4D0CB;align-self:center;}
.mm-sh[data-align="가운데"] .mm-sh__line{width:48px;flex:0 0 auto;margin-top:6px;}
@media (prefers-reduced-motion:no-preference){
  .mm-sh>*{opacity:0;transform:translateY(16px);transition:opacity .6s cubic-bezier(.4,0,.2,1),transform .6s cubic-bezier(.4,0,.2,1);}
  .mm-sh.is-in>*{opacity:1;transform:none;}
  .mm-sh.is-in .mm-sh__text{transition-delay:.08s;}
  .mm-sh.is-in .mm-sh__line{transition-delay:.16s;}
}

/* 여러 줄(text-editor) 입력 시 문단 줄간격 통일 */
.mm-sh__title p{margin:0 0 .15em;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-sh{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function isLeft(al){al=String(al||'').trim();var low=al.toLowerCase();return al==='true'||al.indexOf('왼')>=0||al.indexOf('좌')>=0||low.indexOf('left')>=0||low.indexOf('start')>=0;}
  function init(){
    var l=document.querySelectorAll('.mm-sh');
    if(!l.length){return setTimeout(init,50);}
    for(var k=0;k<l.length;k++) l[k].setAttribute('data-align', isLeft(l[k].getAttribute('data-align'))?'왼쪽':'가운데');
    if(!('IntersectionObserver' in window)){for(var j=0;j<l.length;j++)l[j].classList.add('is-in');return;}
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('is-in');io.unobserve(e.target);}});},{threshold:.3});
    for(var i=0;i<l.length;i++)io.observe(l[i]);
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## N8. 마그네틱 CTA 버튼
`폴더: n08_magnetic_cta`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "마우스 따라 끌려오는 CTA 버튼(PC)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 마그네틱 CTA 버튼
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 마우스를 따라 미세하게 끌려오는 프리미엄 CTA (주목·전환)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name text @type outlined-textfield @default "간편 진단 시작하기" @label "버튼 문구" --}}
{{!-- @name link @type outlined-textfield @default "" @label "링크" --}}
{{!-- @name theme @type outlined-textfield @default "다크" @label "테마 — 입력: 다크 · 라이트" --}}
{{!-- @name hint @type outlined-textfield @default "" @label "버튼 아래 작은 문구(선택)" --}}
<div class="mm-mag" data-theme="{{theme}}">
  <a class="mm-mag__btn" href="{{link}}"><span class="mm-mag__label">{{text}}</span><span class="mm-mag__arrow" aria-hidden="true">→</span></a>
  <p class="mm-mag__hint">{{hint}}</p>
</div>
```
### CSS 탭
```css
.mm-mag{text-align:center;padding:clamp(20px,4vw,32px) 12px;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-mag__btn{display:inline-flex;align-items:center;gap:10px;padding:18px 38px;border-radius:999px;background:#1A1A1A;color:#FAF9F7;font-size:clamp(15px,4vw,17px);font-weight:700;text-decoration:none;will-change:transform;transition:transform .25s cubic-bezier(.2,.8,.2,1),box-shadow .25s,background .2s;box-shadow:0 6px 24px rgba(26,26,26,.18);}
.mm-mag[data-theme="라이트"] .mm-mag__btn{background:#FAF9F7;color:#1A1A1A;border:1px solid #1A1A1A;box-shadow:0 6px 24px rgba(26,26,26,.1);}
.mm-mag__arrow{display:inline-block;transition:transform .25s cubic-bezier(.2,.8,.2,1);}
.mm-mag__btn:hover .mm-mag__arrow{transform:translateX(4px);}
.mm-mag__hint{margin:14px 0 0;font-size:12px;color:#8A8580;}
.mm-mag__hint:empty{display:none;}
@media (prefers-reduced-motion:reduce){.mm-mag__btn{transition:none;}}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-mag{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function bind(btn){
    var STR=0.35,MAX=14;
    btn.addEventListener('pointermove',function(e){
      if(e.pointerType==='touch')return;
      var r=btn.getBoundingClientRect();
      var x=(e.clientX-(r.left+r.width/2))*STR;
      var y=(e.clientY-(r.top+r.height/2))*STR;
      x=Math.max(-MAX,Math.min(MAX,x));y=Math.max(-MAX,Math.min(MAX,y));
      btn.style.transform='translate('+x+'px,'+y+'px)';
    });
    btn.addEventListener('pointerleave',function(){btn.style.transform='';});
  }
  function init(){var l=document.querySelectorAll('.mm-mag__btn');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)bind(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## N9. 신념·수치 그리드
`폴더: n09_stat_grid`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "큰 글자/숫자 + 작은 라벨을 그리드로(신념·100%·ZERO 등). 숫자는 카운트업." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 신념·수치 그리드
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 큰 글자(신념·ZERO) 또는 숫자(100%·1240) + 작은 라벨을 2열 그리드로
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name theme @type outlined-textfield @default "다크" @label "테마 — 입력: 다크 · 라이트" --}}
{{!-- @name maxWidth @type outlined-textfield @default "760" @label "가로 최대폭 — 숫자 자유 입력: 예 1280 (비우면 760, full=꽉차게). 넓혀도 항목은 중앙 1행, 검은 영역만 넓어짐" --}}
{{!-- @name items @type item @label "항목" --}}
<div class="mm-belief" data-maxw="{{maxWidth}}" data-theme="{{theme}}">
{{#each items}}
  {{!-- @name big @type outlined-textfield @default "신념" @label "큰 글자/숫자 (예: 신념, 100, ZERO)" --}}
  {{!-- @name suffix @type outlined-textfield @default "" @label "단위(선택, 예: %)" --}}
  {{!-- @name label @type outlined-textfield @default "작은 설명" @label "작은 라벨" --}}
  <div class="mm-belief__item">
    <span class="mm-belief__big" data-suffix="{{suffix}}">{{big}}</span>
    <span class="mm-belief__label">{{label}}</span>
  </div>
{{/each}}
</div>
```
### CSS 탭
```css
.mm-belief{box-sizing:border-box;max-width:760px;margin:0 auto;padding:clamp(28px,6vw,48px) clamp(16px,4vw,32px);display:grid;grid-template-columns:repeat(2,1fr);gap:clamp(24px,5vw,40px) clamp(16px,4vw,32px);text-align:center;
  font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;background:#1A1A1A;color:#FAF9F7;}
.mm-belief[data-theme="라이트"]{background:#FAF9F7;color:#1A1A1A;}
/* 가로 최대폭은 자유 숫자값(예 1280) → JS가 max-width 적용(margin auto로 중앙). 검은 영역이 이 폭까지 넓어지고 항목은 중앙에 모임 */
.mm-belief__item{display:flex;flex-direction:column;align-items:center;gap:8px;}
.mm-belief__big{font-family:'Outfit','Plus Jakarta Sans','Pretendard','Noto Sans KR',sans-serif;font-size:clamp(30px,9vw,52px);font-weight:900;line-height:1;letter-spacing:-.03em;color:#FAF9F7;font-variant-numeric:tabular-nums;}
.mm-belief[data-theme="라이트"] .mm-belief__big{color:#1A1A1A;}
/* 한글은 같은 900이어도 영문(Outfit)보다 얇게 보임 → 한글 포함 항목(.is-ko, JS가 부여)만 획을 덧대 두께 보정. 영문·숫자는 그대로 */
.mm-belief__big.is-ko{-webkit-text-stroke:.02em currentColor;}
.mm-belief__label{font-size:clamp(12px,3.2vw,14px);font-weight:600;color:#8A8580;letter-spacing:.01em;}
/* PC: 4개 한 줄. 항목 열폭은 최대 180px로 고정+중앙정렬 → 가로폭을 넓혀도 벌어지지 않고 좌우 여백만 늘어남 */
@media (min-width:720px){.mm-belief{grid-template-columns:repeat(4,minmax(140px,180px));justify-content:center;}}
@media (prefers-reduced-motion:no-preference){
  .mm-belief__item{opacity:0;transform:translateY(18px);transition:opacity .6s cubic-bezier(.4,0,.2,1),transform .6s cubic-bezier(.4,0,.2,1);}
  .mm-belief.is-in .mm-belief__item{opacity:1;transform:none;}
  .mm-belief.is-in .mm-belief__item:nth-child(2){transition-delay:.07s;}
  .mm-belief.is-in .mm-belief__item:nth-child(3){transition-delay:.14s;}
  .mm-belief.is-in .mm-belief__item:nth-child(4){transition-delay:.21s;}
}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-belief{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function fmt(n){return n.toLocaleString('en-US');}
  /* 가로 최대폭: 자유 숫자값(1280 등) 적용. 비움 → 기본(760), full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  function run(root){
    var bigs=root.querySelectorAll('.mm-belief__big');
    for(var i=0;i<bigs.length;i++)(function(el){
      var raw=(el.textContent||'').trim();
      var suf=el.getAttribute('data-suffix')||'';
      if(/[가-힣]/.test(raw)) el.classList.add('is-ko'); // 한글 포함 → 두께 보정 클래스
      else el.classList.remove('is-ko');
      if(/^[\d,]+$/.test(raw)){ // 숫자 → 카운트업
        var target=parseInt(raw.replace(/,/g,''),10)||0,t0=null,dur=1300;
        function step(ts){if(!t0)t0=ts;var p=Math.min((ts-t0)/dur,1),e=1-Math.pow(1-p,3);el.textContent=fmt(Math.round(target*e))+suf;if(p<1)requestAnimationFrame(step);}
        requestAnimationFrame(step);
      } else { el.textContent=raw+suf; } // 글자 → 그대로(+단위)
    })(bigs[i]);
  }
  function initOne(root){
    /* 가로 최대폭 적용 + 패널값 바뀌면 즉시 재적용 → 편집기 실시간 반영 */
    applyMaxw(root);
    if('MutationObserver' in window){ new MutationObserver(function(){applyMaxw(root);}).observe(root,{attributes:true,attributeFilter:['data-maxw']}); }
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){root.classList.add('is-in');run(root);io.disconnect();}});},{threshold:.3});
      io.observe(root);
    }else{root.classList.add('is-in');run(root);}
  }
  function init(){var l=document.querySelectorAll('.mm-belief');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## N10. 아이콘 칩 내비
`폴더: n10_icon_nav`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "라벨+링크 칩 내비. 기본 중앙, 넘치면 가로 스크롤(모바일). PC/모바일 글씨·칩 높이 조절." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 칩 내비
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 라벨 알약 버튼들. 중앙정렬·모바일 가로스크롤(우측 peek)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name theme @type outlined-textfield @default "다크" @label "테마 — 입력: 다크 · 라이트" --}}
{{!-- @name align @type outlined-textfield @default "가운데" @label "기본 정렬 — 입력: 가운데 · 왼쪽 (칩이 넘치면 자동 가로 스크롤)" --}}
{{!-- @name fontPC @type outlined-textfield @default "11" @label "PC 글씨 크기(pt)" --}}
{{!-- @name fontMobile @type outlined-textfield @default "10" @label "모바일 글씨 크기(pt)" --}}
{{!-- @name padY @type outlined-textfield @default "12" @label "칩 상하 여백(px) — 칩 높이 조절" --}}
{{!-- @name chips @type item @label "칩" --}}
<div class="mm-nav" data-theme="{{theme}}" data-align="{{align}}" data-fpc="{{fontPC}}" data-fm="{{fontMobile}}" data-pady="{{padY}}">
  <div class="mm-nav__track">
  {{#each chips}}
    {{!-- @name label @type outlined-textfield @default "메뉴" @label "라벨" --}}
    {{!-- @name link @type outlined-textfield @default "" @label "링크" --}}
    <a class="mm-nav__chip" href="{{link}}"><span class="mm-nav__label">{{label}}</span></a>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-nav{--nav-fpc:11pt;--nav-fm:10pt;--nav-pady:12px;
  box-sizing:border-box;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;scroll-snap-type:x proximity;padding:8px 12px;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-nav::-webkit-scrollbar{display:none;}
/* 트랙: 다 들어오면 margin auto로 중앙, 넘치면 좌측부터 스크롤 + 우측 칩 살짝 보임(peek). justify-content:center의 좌측 잘림 없음 */
.mm-nav__track{display:flex;flex-wrap:nowrap;gap:10px;width:max-content;margin:0 auto;}
.mm-nav[data-align="왼쪽"] .mm-nav__track{margin:0;}
/* 칩 상하 여백=--nav-pady, 폰트=PC/모바일 각각(pt), 아이콘은 폰트에 비례(em) */
.mm-nav__chip{flex:0 0 auto;scroll-snap-align:start;display:inline-flex;align-items:center;padding:var(--nav-pady) 20px;border-radius:999px;text-decoration:none;font-size:var(--nav-fpc);line-height:1;font-weight:700;letter-spacing:-.01em;transition:transform .2s cubic-bezier(.4,0,.2,1),box-shadow .25s,opacity .2s;}
@media (max-width:768px){.mm-nav__chip{font-size:var(--nav-fm);}}
.mm-nav[data-theme="다크"] .mm-nav__chip{background:#1A1A1A;color:#FAF9F7;}
.mm-nav[data-theme="라이트"] .mm-nav__chip{background:#FFFFFF;color:#1A1A1A;border:1px solid #D4D0CB;}
.mm-nav__chip:active{transform:scale(.97);}
@media (hover:hover){
  .mm-nav[data-theme="다크"] .mm-nav__chip:hover{opacity:.88;transform:translateY(-1px);}
  .mm-nav[data-theme="라이트"] .mm-nav__chip:hover{border-color:#1A1A1A;transform:translateY(-1px);box-shadow:0 4px 14px rgba(0,0,0,.06);}
}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-nav{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  /* 숫자값 → 단위 부여(폰트 pt / 여백 px) 후 CSS 변수 주입 */
  function unit(v,u){v=String(v==null?'':v).trim(); if(!v)return null; if(String(parseFloat(v))===v)v+=u; return v;}
  function applySettings(root){
    var fpc=unit(root.getAttribute('data-fpc'),'pt'); if(fpc)root.style.setProperty('--nav-fpc',fpc);
    var fm=unit(root.getAttribute('data-fm'),'pt'); if(fm)root.style.setProperty('--nav-fm',fm);
    var py=unit(root.getAttribute('data-pady'),'px'); if(py)root.style.setProperty('--nav-pady',py);
  }
  function initOne(root){
    applySettings(root);
    /* 폰트·여백 값 바뀌면 즉시 반영(편집기 실시간). 정렬(가운데/왼쪽)은 CSS 속성선택자가 담당 */
    if('MutationObserver' in window){ new MutationObserver(function(){applySettings(root);}).observe(root,{attributes:true,attributeFilter:['data-fpc','data-fm','data-pady']}); }
  }
  function init(){var l=document.querySelectorAll('.mm-nav');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## N11. 유튜브 썸네일 배너
`폴더: n11_youtube_banner`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "유튜브 영상 썸네일+재생버튼. 클릭하면 유튜브로 열림(인라인 재생은 iframe금지라 불가)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 유튜브 썸네일 배너
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 영상 URL만 넣으면 썸네일 자동 + 재생버튼 → 클릭 시 유튜브로 이동
  🚫 fetch·iframe 0 (썸네일은 이미지 로드, 인라인 재생 아님)
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "" @label "제목(선택)" --}}
{{!-- @name colsPC @type outlined-textfield @default "3" @label "PC 한 줄 개수 — 예: 2 · 3 · 4" --}}
{{!-- @name maxw @type outlined-textfield @default "" @label "가로 최대폭 — 숫자 자유 입력(예 1280 · 비우면 꽉 채움). 영역 확장해도 이 값에서 멈춤" --}}
{{!-- @name videos @type item @label "영상" --}}
<div class="mm-yt" data-cols="{{colsPC}}" data-maxw="{{maxw}}">
  <p class="mm-yt__heading">{{title}}</p>
  <div class="mm-yt__row">
  {{#each videos}}
    {{!-- @name url @type outlined-textfield @default "" @label "유튜브 URL 또는 영상ID" --}}
    {{!-- @name caption @type outlined-textfield @default "" @label "영상 제목(선택)" --}}
    {{!-- @name thumb @type image @label "썸네일 직접 지정(선택) (권장 1600×900px)" --}}
    <a class="mm-yt__card" href="https://www.youtube.com" target="_blank" rel="noopener" data-url="{{url}}" data-thumb="{{thumb}}">
      <span class="mm-yt__thumb">
        <img class="mm-yt__img" src="" alt="" loading="lazy">
        <span class="mm-yt__play" aria-hidden="true"></span>
      </span>
      <span class="mm-yt__cap">{{caption}}</span>
    </a>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-yt{--yt-cols:3;box-sizing:border-box;max-width:920px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-yt__heading{margin:0 0 16px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(20px,5vw,28px);font-weight:800;letter-spacing:-.02em;}
.mm-yt__heading:empty{display:none;}
/* PC: 한 줄 N개(--yt-cols) 그리드 */
.mm-yt__row{display:grid;grid-template-columns:repeat(var(--yt-cols),minmax(0,1fr));gap:16px;}
/* 모바일: 1행 가로 스크롤 + 우측 카드 살짝 보임(peek) */
@media (max-width:768px){
  .mm-yt__row{display:flex;flex-wrap:nowrap;gap:12px;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;scroll-snap-type:x proximity;padding-bottom:4px;}
  .mm-yt__row::-webkit-scrollbar{display:none;}
  .mm-yt__card{flex:0 0 78%;scroll-snap-align:start;}
}
.mm-yt__card{display:block;text-decoration:none;color:inherit;}
.mm-yt__thumb{position:relative;display:block;aspect-ratio:16/9;border-radius:12px;overflow:hidden;background:#F5F3F0;}
.mm-yt__img{width:100%;height:100%;object-fit:cover;display:block;transition:transform .4s cubic-bezier(.4,0,.2,1);}
.mm-yt__play{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);width:54px;height:54px;border-radius:50%;background:rgba(26,26,26,.72);box-shadow:0 4px 16px rgba(0,0,0,.25);transition:background .2s,transform .2s;}
.mm-yt__play::after{content:'';position:absolute;top:50%;left:54%;transform:translate(-50%,-50%);width:0;height:0;border-left:15px solid #FAF9F7;border-top:9px solid transparent;border-bottom:9px solid transparent;}
.mm-yt__cap{display:block;margin-top:10px;font-size:clamp(13px,3.6vw,15px);font-weight:600;line-height:1.45;color:#2D2D2D;}
.mm-yt__cap:empty{display:none;}
@media (hover:hover){.mm-yt__card:hover .mm-yt__img{transform:scale(1.05);}.mm-yt__card:hover .mm-yt__play{background:#1A1A1A;transform:translate(-50%,-50%) scale(1.06);}}

/* 모바일 좌측 여백만(가로 스크롤이라 우측은 카드 peek 유지) */
@media (max-width:768px){.mm-yt{box-sizing:border-box;padding-left:16px;}}
```
### JS 탭
```js
(function(){
  function vid(u){
    u=String(u||'').trim();
    if(/^[\w-]{11}$/.test(u))return u; // 그냥 ID
    var m=u.match(/(?:youtu\.be\/|v=|\/embed\/|\/shorts\/)([\w-]{11})/);
    return m?m[1]:'';
  }
  /* 가로 최대폭: 자유 숫자값(1280 등). 비움/full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  function applySettings(root){
    var c=String(root.getAttribute('data-cols')||'').trim();
    if(/^[1-9]\d*$/.test(c)) root.style.setProperty('--yt-cols',c);
    applyMaxw(root);
  }
  function initOne(root){
    applySettings(root);
    /* PC 열수·최대폭 바뀌면 즉시 반영(편집기 실시간) */
    if('MutationObserver' in window){ new MutationObserver(function(){applySettings(root);}).observe(root,{attributes:true,attributeFilter:['data-cols','data-maxw']}); }
    var cards=root.querySelectorAll('.mm-yt__card');
    for(var i=0;i<cards.length;i++){
      var id=vid(cards[i].getAttribute('data-url'));
      var custom=cards[i].getAttribute('data-thumb');
      var img=cards[i].querySelector('.mm-yt__img');
      if(id){
        cards[i].setAttribute('href','https://www.youtube.com/watch?v='+id);
        if(img)img.src= (custom&&custom.length>4) ? custom : 'https://img.youtube.com/vi/'+id+'/hqdefault.jpg';
      } else if(custom&&custom.length>4&&img){ img.src=custom; }
    }
  }
  function init(){var l=document.querySelectorAll('.mm-yt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## N12. 카테고리 이미지 그리드 (트렌디)
`폴더: n12_category_grid`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "카테고리 이미지 타일 1행 가로 진열. PC 중앙, 모바일 가로 스크롤(우측 타일 peek). PC/모바일 이미지·높이 각각." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 카테고리 이미지 1행 진열
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 사진 타일 1행 가로. PC 중앙정렬, 모바일 가로스크롤(우측 peek). PC/모바일 사진·높이 각각
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name align @type outlined-textfield @default "가운데" @label "정렬 — 입력: 가운데 · 왼쪽 (넘치면 자동 가로 스크롤)" --}}
{{!-- @name heightPC @type outlined-textfield @default "300" @label "PC 타일 높이(px)" --}}
{{!-- @name heightMobile @type outlined-textfield @default "220" @label "모바일 타일 높이(px)" --}}
{{!-- @name maxw @type outlined-textfield @default "" @label "가로 최대폭 — 숫자 자유 입력(예 1280 · 비우면 꽉 채움). 영역 확장해도 이 값에서 멈춤" --}}
{{!-- @name tiles @type item @label "카테고리 타일" --}}
<div class="mm-cat" data-align="{{align}}" data-hpc="{{heightPC}}" data-hm="{{heightMobile}}" data-maxw="{{maxw}}">
  <div class="mm-cat__track">
  {{#each tiles}}
    {{!-- @name image @type image @label "PC 이미지 — 가로형 권장(높이에 맞춰 표시)" --}}
    {{!-- @name imageMobile @type image @label "모바일 이미지 (선택) — 세로형 가능 · 비우면 PC 이미지 사용" --}}
    {{!-- @name label @type outlined-textfield @default "카테고리" @label "큰 라벨" --}}
    {{!-- @name sublabel @type outlined-textfield @default "" @label "작은 설명(선택)" --}}
    {{!-- @name link @type outlined-textfield @default "" @label "링크" --}}
    <a class="mm-cat__tile" href="{{link}}">
      <img class="mm-cat__img" src="{{image}}" alt="">
      <img class="mm-cat__imgm" src="{{imageMobile}}" alt="">
      <span class="mm-cat__veil" aria-hidden="true"></span>
      <span class="mm-cat__cap">
        <span class="mm-cat__label">{{label}}</span>
        <span class="mm-cat__sub">{{sublabel}}</span>
      </span>
      <span class="mm-cat__arrow" aria-hidden="true">→</span>
    </a>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-cat{--cat-h:300px;--cat-hm:220px;box-sizing:border-box;margin-left:auto;margin-right:auto;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;scroll-snap-type:x proximity;padding:4px;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-cat::-webkit-scrollbar{display:none;}
/* 트랙: 다 들어오면 margin auto로 중앙(PC), 넘치면 좌측부터 스크롤+우측 타일 살짝 보임(peek, 모바일) */
.mm-cat__track{display:flex;flex-wrap:nowrap;gap:12px;width:max-content;margin:0 auto;}
.mm-cat[data-align="왼쪽"] .mm-cat__track{margin:0;}
/* 타일: 높이 고정(PC=--cat-h / 모바일=--cat-hm), 너비=이미지 비율 자동. flex 아이템이라 이미지 폭으로 shrink-wrap */
.mm-cat__tile{flex:0 0 auto;scroll-snap-align:start;position:relative;height:var(--cat-h);overflow:hidden;border-radius:16px;background:#1A1A1A;text-decoration:none;}
@media (max-width:768px){.mm-cat__tile{height:var(--cat-hm);}}
.mm-cat__img,.mm-cat__imgm{height:100%;width:auto;max-width:none;display:block;transition:transform .55s cubic-bezier(.4,0,.2,1);}
.mm-cat__imgm{display:none;}
/* 모바일 전용 이미지가 있으면(data-hasm=1) 모바일에서 스왑 */
@media (max-width:768px){
  .mm-cat__tile[data-hasm="1"] .mm-cat__img{display:none;}
  .mm-cat__tile[data-hasm="1"] .mm-cat__imgm{display:block;}
}
.mm-cat__veil{position:absolute;inset:0;background:linear-gradient(rgba(26,26,26,0) 38%,rgba(26,26,26,.8));}
.mm-cat__cap{position:absolute;left:0;right:0;bottom:0;padding:clamp(14px,2.5vw,22px);color:#FAF9F7;display:flex;flex-direction:column;gap:3px;}
.mm-cat__label{font-family:'Outfit','Plus Jakarta Sans','Pretendard','Noto Sans KR',sans-serif;font-size:clamp(16px,2.6vw,21px);font-weight:800;letter-spacing:-.02em;}
.mm-cat__sub{font-size:clamp(12px,1.8vw,13px);color:#D4D0CB;font-weight:600;}
.mm-cat__sub:empty{display:none;}
.mm-cat__arrow{position:absolute;top:14px;right:14px;width:34px;height:34px;border-radius:50%;background:rgba(250,249,247,.16);display:flex;align-items:center;justify-content:center;font-size:16px;color:#FAF9F7;transition:background .3s,transform .3s;}
@media (hover:hover){
  .mm-cat__tile:hover .mm-cat__img,.mm-cat__tile:hover .mm-cat__imgm{transform:scale(1.07);}
  .mm-cat__tile:hover .mm-cat__arrow{background:#FAF9F7;color:#1A1A1A;transform:translateX(2px);}
}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-cat{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function unit(v,u){v=String(v==null?'':v).trim(); if(!v)return null; if(String(parseFloat(v))===v)v+=u; return v;}
  /* 가로 최대폭: 자유 숫자값(1280 등). 비움/full → 꽉 채움 */
  function applyMaxw(root){
    var mw=root.getAttribute('data-maxw'); mw=(mw==null?'':mw).trim(); var low=mw.toLowerCase();
    if(!mw){ root.style.maxWidth=''; }
    else if(low==='full'||low==='none'){ root.style.maxWidth='none'; }
    else { if(String(parseFloat(mw))===mw) mw+='px'; root.style.maxWidth=mw; }
  }
  /* PC/모바일 타일 높이 → CSS 변수 주입 */
  function applySettings(root){
    var hpc=unit(root.getAttribute('data-hpc'),'px'); if(hpc)root.style.setProperty('--cat-h',hpc);
    var hm=unit(root.getAttribute('data-hm'),'px'); if(hm)root.style.setProperty('--cat-hm',hm);
    applyMaxw(root);
  }
  function initOne(root){
    applySettings(root);
    /* 타일별 모바일 이미지가 실제 있으면 data-hasm=1 → 모바일에서 스왑 */
    var tiles=root.querySelectorAll('.mm-cat__tile');
    for(var i=0;i<tiles.length;i++){
      var m=tiles[i].querySelector('.mm-cat__imgm');
      if(m && String(m.getAttribute('src')||'').trim()) tiles[i].setAttribute('data-hasm','1');
    }
    /* 높이·최대폭 값 바뀌면 즉시 반영(편집기 실시간). 정렬은 CSS 속성선택자가 담당 */
    if('MutationObserver' in window){ new MutationObserver(function(){applySettings(root);}).observe(root,{attributes:true,attributeFilter:['data-hpc','data-hm','data-maxw']}); }
  }
  function init(){var l=document.querySelectorAll('.mm-cat');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## T1. 스크롤 스토리텔링 (스티키)
`폴더: t01_storytelling`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "스크롤에 따라 장면·이미지 전환(PC 스티키)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 스크롤 스토리텔링(스티키)
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 이미지 고정 + 스크롤에 따라 장면(문구/이미지) 전환 (브랜드 몰입)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name scenes @type item @label "장면" --}}
<div class="mm-st">
  <div class="mm-st__media"><img class="mm-st__mediaimg" src="" alt=""></div>
  <div class="mm-st__scenes">
  {{#each scenes}}
    {{!-- @name image @type image @label "이미지 — 권장 1200×800px" --}}
    {{!-- @name heading @type outlined-textfield @default "장면 제목" @label "제목" --}}
    {{!-- @name text @type text-editor @default "<p>내용</p>" @label "내용" --}}
    <section class="mm-st__scene" data-img="{{image}}">
      <img class="mm-st__inline" src="{{image}}" alt="">
      <h3 class="mm-st__heading">{{heading}}</h3>
      <div class="mm-st__text">{{text}}</div>
    </section>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-st{max-width:1000px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-st__media{display:none;}
.mm-st__scene{padding:clamp(40px,10vw,90px) 4px;}
.mm-st__inline{width:100%;aspect-ratio:4/3;object-fit:cover;border-radius:14px;background:#F5F3F0;margin-bottom:18px;display:block;}
.mm-st__heading{margin:0 0 12px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(22px,6vw,34px);font-weight:900;letter-spacing:-.02em;line-height:1.2;}
.mm-st__text{font-size:clamp(15px,4vw,17px);line-height:1.7;color:#4A4A4A;}
.mm-st__text p{margin:0 0 .15em;}
@media (min-width:840px){
  .mm-st{display:grid;grid-template-columns:1fr 1fr;gap:48px;align-items:start;}
  .mm-st__media{display:block;position:sticky;top:10vh;height:80vh;}
  .mm-st__mediaimg{width:100%;height:100%;object-fit:cover;border-radius:16px;background:#F5F3F0;transition:opacity .4s cubic-bezier(.4,0,.2,1);}
  .mm-st__inline{display:none;}
  .mm-st__scene{min-height:80vh;display:flex;flex-direction:column;justify-content:center;padding:0;}
  .mm-st__scene{opacity:.35;transition:opacity .4s;}
  .mm-st__scene.is-active{opacity:1;}
}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-st{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var media=root.querySelector('.mm-st__mediaimg');
    var scenes=root.querySelectorAll('.mm-st__scene');
    if(!scenes.length)return;
    if(media&&scenes[0])media.src=scenes[0].getAttribute('data-img')||'';
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){
        for(var k=0;k<scenes.length;k++)scenes[k].classList.toggle('is-active',scenes[k]===e.target);
        if(media){var img=e.target.getAttribute('data-img')||'';if(img&&media.src!==img){media.style.opacity='0';setTimeout(function(){media.src=img;media.style.opacity='1';},200);}}
      }});},{threshold:.5});
      for(var i=0;i<scenes.length;i++)io.observe(scenes[i]);
    }else{scenes[0].classList.add('is-active');}
  }
  function init(){var l=document.querySelectorAll('.mm-st');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## T2. 복원 인터랙티브 타임라인
`폴더: t02_repair_timeline`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "복원 단계 번호·선 타임라인, 스크롤 순차 등장." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 복원 인터랙티브 타임라인
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 수거→검수→수리→출고 단계를 번호·선으로. 스크롤 시 순차 등장(과정 투명성=신뢰)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "복원 과정" @label "제목(선택)" --}}
{{!-- @name steps @type item @label "단계" --}}
<div class="mm-tl">
  <p class="mm-tl__title">{{title}}</p>
  <ol class="mm-tl__list">
  {{#each steps}}
    {{!-- @name stepName @type outlined-textfield @default "단계명" @label "단계명" --}}
    {{!-- @name stepDesc @type text-editor @default "<p>설명</p>" @label "설명" --}}
    {{!-- @name stepImage @type image @label "이미지(선택) — 권장 800×800px" --}}
    <li class="mm-tl__step">
      <span class="mm-tl__idx" aria-hidden="true"></span>
      <div class="mm-tl__body">
        <h3 class="mm-tl__name">{{stepName}}</h3>
        <div class="mm-tl__desc">{{stepDesc}}</div>
        <img class="mm-tl__img" src="{{stepImage}}" alt="">
      </div>
    </li>
  {{/each}}
  </ol>
</div>
```
### CSS 탭
```css
.mm-tl{max-width:680px;margin:0 auto;padding:clamp(20px,4vw,32px) 20px;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-tl__title{margin:0 0 clamp(20px,4vw,32px);font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(20px,5vw,28px);font-weight:800;letter-spacing:-.02em;}
.mm-tl__title:empty{display:none;}
.mm-tl__list{counter-reset:tl;list-style:none;margin:0;padding:0;}
.mm-tl__step{position:relative;display:grid;grid-template-columns:clamp(48px,12vw,72px) 1fr;gap:clamp(14px,3vw,22px);counter-increment:tl;padding-bottom:clamp(28px,6vw,44px);}
.mm-tl__step::before{content:'';position:absolute;left:clamp(23px,5.6vw,35px);top:clamp(36px,9vw,52px);bottom:0;width:1px;background:#D4D0CB;}
.mm-tl__step:last-child::before{display:none;}
.mm-tl__idx{display:flex;align-items:flex-start;justify-content:center;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(24px,6.5vw,38px);font-weight:900;line-height:1;letter-spacing:-.04em;color:#1A1A1A;}
.mm-tl__idx::before{content:counter(tl,decimal-leading-zero);}
.mm-tl__name{margin:0 0 8px;font-size:clamp(16px,4.2vw,20px);font-weight:800;letter-spacing:-.01em;}
.mm-tl__desc{font-size:clamp(14px,3.8vw,15px);line-height:1.65;color:#4A4A4A;}
.mm-tl__desc p{margin:0 0 .15em;}
.mm-tl__img{width:100%;max-width:360px;aspect-ratio:16/10;object-fit:cover;border-radius:10px;background:#F5F3F0;margin-top:14px;display:block;}
@media (prefers-reduced-motion:no-preference){
  .mm-tl__step{opacity:0;transform:translateY(22px);transition:opacity .55s cubic-bezier(.4,0,.2,1),transform .55s cubic-bezier(.4,0,.2,1);}
  .mm-tl__step.is-in{opacity:1;transform:none;}
}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-tl{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var steps=root.querySelectorAll('.mm-tl__step');
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('is-in');io.unobserve(e.target);}});},{threshold:.2,rootMargin:'0px 0px -10% 0px'});
      for(var i=0;i<steps.length;i++)io.observe(steps[i]);
    }else{for(var j=0;j<steps.length;j++)steps[j].classList.add('is-in');}
  }
  function init(){var l=document.querySelectorAll('.mm-tl');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## T3. 가위 해부도 핫스팟
`폴더: t03_hotspot`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "가위 사진 위 점 클릭 시 부위 설명 툴팁." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 가위 해부도 핫스팟
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 가위 사진 위 점을 클릭 → 부위 설명 툴팁 (전문성·교육)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name baseImage @type image @label "가위 사진(배경) (권장 1600×900px)" --}}
{{!-- @name ratio @type outlined-textfield @default "16/9" @label "비율 — 입력: 16/9 · 4/3 · 1/1" --}}
{{!-- @name spots @type item @label "핫스팟" --}}
<div class="mm-hs">
  <div class="mm-hs__stage" data-ar="{{ratio}}">
    <img class="mm-hs__base" src="{{baseImage}}" alt="" draggable="false">
    {{#each spots}}
      {{!-- @name x @type outlined-textfield @default "50" @label "가로 위치 %(0~100)" --}}
      {{!-- @name y @type outlined-textfield @default "50" @label "세로 위치 %(0~100)" --}}
      {{!-- @name spotTitle @type outlined-textfield @default "부위" @label "부위명" --}}
      {{!-- @name spotDesc @type text-editor @default "<p>설명</p>" @label "설명" --}}
      <div class="mm-hs__spot" data-x="{{x}}" data-y="{{y}}">
        <button type="button" class="mm-hs__dot" aria-label="{{spotTitle}}"></button>
        <div class="mm-hs__tip"><strong>{{spotTitle}}</strong><div class="mm-hs__tipdesc">{{spotDesc}}</div></div>
      </div>
    {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-hs{max-width:760px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-hs__stage{position:relative;width:100%;aspect-ratio:16/9;overflow:hidden;border-radius:14px;background:#F5F3F0;}
.mm-hs__base{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;display:block;}
.mm-hs__spot{position:absolute;transform:translate(-50%,-50%);z-index:2;}
.mm-hs__dot{width:22px;height:22px;border-radius:50%;border:2px solid #FAF9F7;background:#1A1A1A;cursor:pointer;padding:0;box-shadow:0 2px 8px rgba(0,0,0,.3);position:relative;}
.mm-hs__dot::after{content:'';position:absolute;inset:-7px;border-radius:50%;border:1px solid rgba(250,249,247,.6);animation:mmhs-pulse 2s ease-out infinite;}
@keyframes mmhs-pulse{0%{transform:scale(.7);opacity:.9}100%{transform:scale(1.5);opacity:0}}
.mm-hs__tip{position:absolute;left:50%;bottom:calc(100% + 12px);transform:translateX(-50%) translateY(6px);width:max-content;max-width:220px;background:#1A1A1A;color:#FAF9F7;padding:12px 14px;border-radius:10px;opacity:0;visibility:hidden;transition:opacity .2s,transform .2s;z-index:3;text-align:left;box-shadow:0 8px 24px rgba(0,0,0,.25);}
.mm-hs__tip::after{content:'';position:absolute;top:100%;left:50%;transform:translateX(-50%);border:6px solid transparent;border-top-color:#1A1A1A;}
.mm-hs__spot.is-open .mm-hs__tip{opacity:1;visibility:visible;transform:translateX(-50%) translateY(0);}
.mm-hs__spot.is-open .mm-hs__dot{background:#FAF9F7;}
.mm-hs__spot.is-open .mm-hs__dot::after{display:none;}
.mm-hs__tip strong{display:block;font-size:14px;font-weight:800;margin-bottom:4px;}
.mm-hs__tipdesc{font-size:13px;line-height:1.55;color:#D4D0CB;}
.mm-hs__tipdesc p{margin:0 0 .15em;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-hs{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var spots=root.querySelectorAll('.mm-hs__spot');
    function closeAll(except){for(var k=0;k<spots.length;k++){if(spots[k]!==except)spots[k].classList.remove('is-open');}}
    for(var i=0;i<spots.length;i++)(function(sp){
      var dot=sp.querySelector('.mm-hs__dot');
      dot.addEventListener('click',function(e){e.stopPropagation();var on=sp.classList.contains('is-open');closeAll(sp);sp.classList.toggle('is-open',!on);});
    })(spots[i]);
    document.addEventListener('click',function(){closeAll(null);});
  }
  function init(){var l=document.querySelectorAll('.mm-hs');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 동적 스타일(비율/배경/좌표): 인라인 style {{}}는 아임웹 저장거부 → data-* 속성을 JS로 적용 */
(function(){function ap(){
var A=document.querySelectorAll("[data-ar]");for(var i=0;i<A.length;i++){var v=(A[i].getAttribute("data-ar")||"").trim();if(v){if(A[i].classList.contains("mm-cat"))A[i].style.setProperty("--mm-ratio",v);else A[i].style.aspectRatio=v;}}
var B=document.querySelectorAll("[data-bg]");for(var i=0;i<B.length;i++){var v=(B[i].getAttribute("data-bg")||"").trim();if(v)B[i].style.backgroundImage="url('"+v+"')";}
var C=document.querySelectorAll("[data-x]");for(var i=0;i<C.length;i++){var x=(C[i].getAttribute("data-x")||"").trim(),y=(C[i].getAttribute("data-y")||"").trim();if(x)C[i].style.left=x+"%";if(y)C[i].style.top=y+"%";}
}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",ap);else ap();})();
```

---

## T4. 3D 틸트 제품 카드
`폴더: t04_tilt_cards`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "마우스 따라 카드가 3D로 기울어짐(PC)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 3D 틸트 제품 카드
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 마우스 따라 카드가 살짝 기울어지는 프리미엄 인터랙션
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name cards @type item @label "카드" --}}
<div class="mm-tilt">
{{#each cards}}
  {{!-- @name image @type image @label "이미지 — 권장 800×1000px (세로 카드)" --}}
  {{!-- @name title @type outlined-textfield @default "제목" @label "제목" --}}
  {{!-- @name desc @type outlined-textfield @default "" @label "설명(선택)" --}}
  {{!-- @name link @type outlined-textfield @default "" @label "링크(선택)" --}}
  <a class="mm-tilt__card" href="{{link}}">
    <div class="mm-tilt__inner">
      <img class="mm-tilt__img" src="{{image}}" alt="">
      <h3 class="mm-tilt__title">{{title}}</h3>
      <p class="mm-tilt__desc">{{desc}}</p>
    </div>
  </a>
{{/each}}
</div>
```
### CSS 탭
```css
.mm-tilt{max-width:920px;margin:0 auto;display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:16px;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-tilt__card{display:block;text-decoration:none;color:inherit;perspective:800px;}
.mm-tilt__inner{background:#FFFFFF;border:1px solid #D4D0CB;border-radius:16px;padding:16px;transition:transform .15s ease,box-shadow .3s ease;transform-style:preserve-3d;will-change:transform;}
@media (hover:hover){.mm-tilt__card:hover .mm-tilt__inner{box-shadow:0 18px 40px rgba(0,0,0,.12);}}
.mm-tilt__img{width:100%;aspect-ratio:4/3;object-fit:cover;border-radius:10px;background:#F5F3F0;margin-bottom:14px;display:block;transform:translateZ(28px);}
.mm-tilt__title{margin:0 0 4px;font-size:clamp(15px,4vw,18px);font-weight:800;letter-spacing:-.01em;transform:translateZ(18px);}
.mm-tilt__desc{margin:0;font-size:13px;color:#8A8580;line-height:1.5;transform:translateZ(12px);}
.mm-tilt__desc:empty{display:none;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-tilt{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function bind(card){
    var inner=card.querySelector('.mm-tilt__inner');
    if(!inner)return;
    var MAX=8;
    card.addEventListener('pointermove',function(e){
      if(e.pointerType==='touch')return;
      var r=card.getBoundingClientRect();
      var px=(e.clientX-r.left)/r.width-.5, py=(e.clientY-r.top)/r.height-.5;
      inner.style.transform='rotateY('+(px*MAX)+'deg) rotateX('+(-py*MAX)+'deg)';
    });
    card.addEventListener('pointerleave',function(){inner.style.transform='';});
  }
  function initOne(root){var cards=root.querySelectorAll('.mm-tilt__card');for(var i=0;i<cards.length;i++)bind(cards[i]);}
  function init(){var l=document.querySelectorAll('.mm-tilt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## T5. 무한 마퀴 띠
`폴더: t05_marquee`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "키워드가 끊임없이 흐르는 띠." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 무한 마퀴 띠
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 핵심 키워드가 끊임없이 흐르는 띠 (트렌디·브랜드 키워드)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name words @type item @label "키워드" --}}
{{!-- @name theme @type outlined-textfield @default "다크" @label "테마 — 입력: 다크 · 라이트" --}}
{{!-- @name speed @type outlined-textfield @default "30" @label "속도(초, 작을수록 빠름)" --}}
{{!-- @name radius @type outlined-textfield @default "12px" @label "모서리 둥글기 — 예: 12px · 0px이면 각지게" --}}
<div class="mm-mq" data-radius="{{radius}}" data-theme="{{theme}}" data-speed="{{speed}}">
  <div class="mm-mq__track">
  {{#each words}}
    {{!-- @name word @type outlined-textfield @default "키워드" @label "키워드" --}}
    <span class="mm-mq__item">{{word}}</span><span class="mm-mq__dot" aria-hidden="true">✦</span>
  {{/each}}
  </div>
</div>
```
### CSS 탭
```css
.mm-mq{overflow:hidden;width:100%;padding:clamp(14px,3vw,22px) 0;background:#1A1A1A;}
.mm-mq[data-theme="라이트"]{background:#FAF9F7;border:1px solid #D4D0CB;}
.mm-mq__track{display:flex;width:max-content;will-change:transform;animation:mm-mq-scroll var(--mq-dur,30s) linear infinite;}
@keyframes mm-mq-scroll{from{transform:translateX(0)}to{transform:translateX(-50%)}}
.mm-mq:hover .mm-mq__track{animation-play-state:paused;}
.mm-mq__item{font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:clamp(16px,4vw,24px);font-weight:800;letter-spacing:-.01em;color:#FAF9F7;white-space:nowrap;padding:0 clamp(14px,3vw,24px);}
.mm-mq[data-theme="라이트"] .mm-mq__item{color:#1A1A1A;}
.mm-mq__dot{color:#8A8580;font-size:clamp(10px,2.4vw,14px);align-self:center;}
@media (prefers-reduced-motion:reduce){.mm-mq__track{animation:none;justify-content:center;flex-wrap:wrap;}}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-mq{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var track=root.querySelector('.mm-mq__track');
    if(!track)return;
    // 끊김 없는 루프: 내용 복제(50% 지점에서 동일)
    track.innerHTML=track.innerHTML+track.innerHTML;
    var sp=parseFloat(root.getAttribute('data-speed'));
    track.style.setProperty('--mq-dur',((isNaN(sp)||sp<=0)?30:sp)+'s');
  }
  function init(){var l=document.querySelectorAll('.mm-mq');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
```

---

## T6. 라이트박스 사례 갤러리
`폴더: t06_lightbox_gallery`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "썸네일 클릭 시 확대 모달(좌우 이동·Esc 닫기)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 라이트박스 사례 갤러리
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 썸네일 그리드 → 클릭 시 확대 모달(순수 DOM, iframe 아님). 복원 사례 몰입.
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name photos @type item @label "사진" --}}
<div class="mm-gal">
  <div class="mm-gal__grid">
  {{#each photos}}
    {{!-- @name image @type image @label "사진 — 권장 800×800px" --}}
    {{!-- @name caption @type outlined-textfield @default "" @label "캡션(선택)" --}}
    <button type="button" class="mm-gal__thumb" data-img="{{image}}" data-cap="{{caption}}">
      <img src="{{image}}" alt="" loading="lazy">
    </button>
  {{/each}}
  </div>
  <div class="mm-gal__lb" hidden aria-modal="true" role="dialog">
    <button type="button" class="mm-gal__close" aria-label="닫기">×</button>
    <button type="button" class="mm-gal__nav mm-gal__prev" aria-label="이전">‹</button>
    <figure class="mm-gal__figure"><img class="mm-gal__lbimg" src="" alt=""><figcaption class="mm-gal__cap"></figcaption></figure>
    <button type="button" class="mm-gal__nav mm-gal__next" aria-label="다음">›</button>
  </div>
</div>
```
### CSS 탭
```css
.mm-gal{max-width:920px;margin:0 auto;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;}
.mm-gal__grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:10px;}
.mm-gal__thumb{padding:0;border:none;border-radius:10px;overflow:hidden;cursor:pointer;background:#F5F3F0;aspect-ratio:1/1;}
.mm-gal__thumb img{width:100%;height:100%;object-fit:cover;display:block;transition:transform .35s cubic-bezier(.4,0,.2,1);}
@media (hover:hover){.mm-gal__thumb:hover img{transform:scale(1.06);}}
.mm-gal__lb{position:fixed;inset:0;z-index:99999;background:rgba(26,26,26,.92);display:flex;align-items:center;justify-content:center;padding:24px;animation:mmgal .2s ease-out;}
@keyframes mmgal{from{opacity:0}to{opacity:1}}
.mm-gal__figure{margin:0;max-width:90vw;max-height:86vh;text-align:center;}
.mm-gal__lbimg{max-width:90vw;max-height:78vh;object-fit:contain;border-radius:8px;display:block;margin:0 auto;}
.mm-gal__cap{margin:12px 0 0;color:#FAF9F7;font-size:14px;font-weight:600;}
.mm-gal__cap:empty{display:none;}
.mm-gal__close{position:absolute;top:16px;right:20px;width:42px;height:42px;border:none;border-radius:50%;background:rgba(255,255,255,.12);color:#FAF9F7;font-size:24px;line-height:1;cursor:pointer;}
.mm-gal__nav{position:absolute;top:50%;transform:translateY(-50%);width:46px;height:46px;border:none;border-radius:50%;background:rgba(255,255,255,.12);color:#FAF9F7;font-size:26px;line-height:1;cursor:pointer;}
.mm-gal__prev{left:16px;}.mm-gal__next{right:16px;}
@media (hover:hover){.mm-gal__close:hover,.mm-gal__nav:hover{background:rgba(255,255,255,.24);}}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-gal{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var thumbs=root.querySelectorAll('.mm-gal__thumb');
    var lb=root.querySelector('.mm-gal__lb'),img=root.querySelector('.mm-gal__lbimg'),cap=root.querySelector('.mm-gal__cap');
    if(!thumbs.length||!lb)return;
    var cur=0;
    function open(i){cur=(i%thumbs.length+thumbs.length)%thumbs.length;img.src=thumbs[cur].getAttribute('data-img')||'';cap.textContent=thumbs[cur].getAttribute('data-cap')||'';lb.hidden=false;document.documentElement.style.overflow='hidden';}
    function close(){lb.hidden=true;document.documentElement.style.overflow='';}
    for(var x=0;x<thumbs.length;x++)(function(idx){thumbs[idx].addEventListener('click',function(){open(idx);});})(x);
    root.querySelector('.mm-gal__close').addEventListener('click',close);
    root.querySelector('.mm-gal__prev').addEventListener('click',function(){open(cur-1);});
    root.querySelector('.mm-gal__next').addEventListener('click',function(){open(cur+1);});
    lb.addEventListener('click',function(e){if(e.target===lb)close();});
    document.addEventListener('keydown',function(e){if(lb.hidden)return;if(e.key==='Escape')close();else if(e.key==='ArrowLeft')open(cur-1);else if(e.key==='ArrowRight')open(cur+1);});
  }
  function init(){var l=document.querySelectorAll('.mm-gal');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
```

---

## T7. 마우스 스포트라이트 히어로
`폴더: t07_spotlight`

### HTML 탭
```html
{{!-- @name widgetInfo @type outlined-textfield @default "어두운 배경에서 커서 주변만 밝아짐(PC)." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 마우스 스포트라이트 히어로
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 어두운 배경에서 커서 주변만 밝아지며 제품이 드러남 (임팩트)
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name bg @type image @label "PC 배경 이미지 — 권장 1600×900px" --}}
{{!-- @name bgMobile @type image @label "모바일 배경 이미지(선택) — 권장 1080×1350px · 비우면 PC 이미지 사용" --}}
{{!-- @name height @type outlined-textfield @default "" @label "높이 — 예: 460px 또는 60vh (비우면 자동)" --}}
{{!-- @name radius @type outlined-textfield @default "16px" @label "모서리 둥글기 — 예: 16px · 0px이면 각지게" --}}
{{!-- @name headline @type text-editor @default "<p>진짜는 가까이서 드러납니다</p>" @label "헤드라인" --}}
{{!-- @name sub @type text-editor @default "<p></p>" @label "서브 문구(선택)" --}}
<div class="mm-sp" data-radius="{{radius}}" data-height="{{height}}">
  <img class="mm-sp__bg" src="{{bg}}" alt="">
  <img class="mm-sp__bgm" src="{{bgMobile}}" alt="">
  <div class="mm-sp__veil" aria-hidden="true"></div>
  <div class="mm-sp__inner">
    <div class="mm-sp__headline" role="heading" aria-level="2">{{headline}}</div>
    <div class="mm-sp__sub">{{sub}}</div>
  </div>
</div>
```
### CSS 탭
```css
.mm-sp{position:relative;overflow:hidden;min-height:clamp(300px,52vw,460px);display:flex;align-items:center;justify-content:center;text-align:center;background:#1A1A1A;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;--mx:50%;--my:50%;}
.mm-sp__bg,.mm-sp__bgm{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;z-index:0;}
.mm-sp__bgm{display:none;}
@media (max-width:768px){.mm-sp[data-hasm="1"] .mm-sp__bg{display:none;}.mm-sp[data-hasm="1"] .mm-sp__bgm{display:block;}}
.mm-sp__veil{position:absolute;inset:0;z-index:1;background:radial-gradient(circle clamp(120px,22vw,220px) at var(--mx) var(--my),rgba(26,26,26,0) 0%,rgba(26,26,26,.55) 45%,rgba(26,26,26,.94) 80%);transition:background .08s linear;}
.mm-sp__inner{position:relative;z-index:2;padding:clamp(28px,6vw,56px);max-width:680px;pointer-events:none;}
.mm-sp__headline{margin:0;font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:clamp(22px,6vw,40px);font-weight:900;line-height:1.25;letter-spacing:-.02em;color:#FAF9F7;text-shadow:0 2px 20px rgba(0,0,0,.4);white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-sp__sub{margin:14px 0 0;font-size:clamp(14px,3.8vw,16px);line-height:1.6;color:#D4D0CB;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;}
.mm-sp__sub:empty{display:none;}
@media (hover:none){.mm-sp__veil{background:linear-gradient(rgba(26,26,26,.25),rgba(26,26,26,.55));}}

/* 여러 줄(text-editor) 입력 시 문단 줄간격 통일 */
.mm-sp__headline p,.mm-sp__sub p{margin:0 0 .15em;}

/* 모바일 좌우 여백(섹션 100% 확장 시 콘텐츠가 화면 끝에 붙지 않게 · 배경은 border-box라 그대로 블리드) */
@media (max-width:768px){.mm-sp{box-sizing:border-box;padding-left:16px;padding-right:16px;}}
```
### JS 탭
```js
(function(){
  function initOne(root){
    var h=root.getAttribute('data-height');
    if(h){ var hv=h.trim(); if(hv){ if(String(parseFloat(hv))===hv) hv+='px'; root.style.minHeight=hv; } }
    var m=root.querySelector('.mm-sp__bgm');
    if(m && String(m.getAttribute('src')||'').trim()) root.setAttribute('data-hasm','1');
    root.addEventListener('pointermove',function(e){
      if(e.pointerType==='touch')return;
      var r=root.getBoundingClientRect();
      root.style.setProperty('--mx',((e.clientX-r.left)/r.width*100)+'%');
      root.style.setProperty('--my',((e.clientY-r.top)/r.height*100)+'%');
    });
  }
  function init(){var l=document.querySelectorAll('.mm-sp');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
```

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
