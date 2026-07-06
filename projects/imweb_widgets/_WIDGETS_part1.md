# 🧩 MAMORU 아임웹 위젯 — Part 1/5 (01_before_after ~ 10_hours_badge)

> 각 위젯=HTML/CSS/JS 3탭. 🚫 삼중괄호·CSS탭 {{변수}}·인라인 on*= 금지. 이 파일 10종.

## 📑 이 파일의 위젯
- 01. 복원 Before/After 슬라이더 ⭐ — `01_before_after`
- 02. 한정세일 카운트다운 — `02_countdown`
- 03. 가위 스펙 비교표 — `03_compare`
- 04. 등급별 견적 계산기 — `04_calculator`
- 05. 정적 누적 카운터 — `05_counter`
- 06. 후기 캐러셀 (수동입력) — `06_review_carousel`
- 07. 가위 관리법 가이드 (아코디언) — `07_care_guide`
- 08. 시술별 가위 선택 가이드 (필터) — `08_treatment_filter`
- 09. 미용가위 용어사전 (검색) — `09_glossary`
- 10. 영업시간 배지 + 카톡 버튼 — `10_hours_badge`

---

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
  <div class="mm-ba__stage" style="aspect-ratio:{{ratio}};border-radius:{{radius}}">
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
.mm-ba__stage{position:relative;width:100%;aspect-ratio:4/3;overflow:hidden;border-radius:12px;background:#F5F3F0;touch-action:none;cursor:ew-resize;user-select:none;}
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
<div class="mm-cd" style="border-radius:{{radius}}" data-theme="{{theme}}" data-date="{{endDate}}" data-time="{{endTime}}">
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
.mm-cd{max-width:520px;margin:0 auto;padding:clamp(24px,5vw,36px) clamp(20px,4vw,32px);border-radius:16px;text-align:center;
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
<div class="mm-calc" style="border-radius:{{radius}}" data-g1n="{{g1n}}" data-g1d="{{g1d}}" data-g2n="{{g2n}}" data-g2d="{{g2d}}" data-g3n="{{g3n}}" data-g3d="{{g3d}}">
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
.mm-calc{max-width:460px;margin:0 auto;padding:clamp(22px,5vw,32px);border:1px solid #D4D0CB;border-radius:16px;background:#FAF9F7;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
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
<div class="mm-hr" style="border-radius:{{radius}}">
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
.mm-hr{max-width:380px;margin:0 auto;padding:clamp(20px,4vw,28px);text-align:center;border:1px solid #D4D0CB;border-radius:16px;background:#FAF9F7;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-hr__badge{display:inline-flex;align-items:center;gap:8px;padding:8px 16px;border-radius:999px;background:#F5F3F0;font-size:14px;font-weight:700;}
.mm-hr__dot{width:8px;height:8px;border-radius:50%;background:#B8B4AF;}
.mm-hr.is-open .mm-hr__dot{background:#1A1A1A;box-shadow:0 0 0 4px rgba(26,26,26,.12);}
.mm-hr.is-open .mm-hr__badge{background:#1A1A1A;color:#FAF9F7;}
.mm-hr__detail{margin:12px 0 0;font-size:13px;color:#8A8580;min-height:18px;}
.mm-hr__btn{display:inline-flex;align-items:center;justify-content:center;margin-top:16px;padding:13px 28px;border-radius:8px;background:#1A1A1A;color:#FAF9F7;font-weight:700;font-size:15px;text-decoration:none;transition:opacity .2s;}
.mm-hr__btn:empty,.mm-hr__btn[href=""]{display:none;}
@media (hover:hover){.mm-hr__btn:hover{opacity:.88;}}
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
```

---
