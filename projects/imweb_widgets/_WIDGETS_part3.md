# 🧩 MAMORU 아임웹 위젯 — Part 3/5 (d05_dual_choice ~ n09_stat_grid)

> 각 위젯=HTML/CSS/JS 3탭. 🚫 삼중괄호 {{{ }}} · CSS탭 {{변수}} · 인라인 style="…{{}}…"(동적값은 data-*+JS) · 인라인 on*= 금지. 이 파일 10종.

## 📑 이 파일의 위젯
- D5. 듀얼 선택 배너 — `d05_dual_choice`
- N1. 가위 길이 시뮬레이터 — `n01_size_ruler`
- N2. 마모루 vs 일반 비교표 — `n02_versus_table`
- N3. 가로 스크롤 핀 갤러리 — `n03_horizontal_pin`
- N4. 텍스트 마스크 헤드라인 — `n04_text_mask`
- N5. 한정 수량 게이지 — `n05_stock_gauge`
- N6. 이용 안내 가로 스텝 — `n06_steps`
- N7. 에디토리얼 섹션 헤더 — `n07_section_header`
- N8. 마그네틱 CTA 버튼 — `n08_magnetic_cta`
- N9. 신념·수치 그리드 — `n09_stat_grid`

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
{{!-- @name maxWidth @type outlined-textfield @default "760" @label "가로 최대폭 — 입력: 760 · 900 · 1000 · 1100 · 1200 · full(꽉차게). 넓혀도 항목은 중앙 1행, 검은 영역만 넓어짐" --}}
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
.mm-belief{max-width:760px;margin:0 auto;padding:clamp(28px,6vw,48px) clamp(16px,4vw,32px);display:grid;grid-template-columns:repeat(2,1fr);gap:clamp(24px,5vw,40px) clamp(16px,4vw,32px);text-align:center;
  font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;background:#1A1A1A;color:#FAF9F7;}
.mm-belief[data-theme="라이트"]{background:#FAF9F7;color:#1A1A1A;}
/* 가로 최대폭 프리셋(CSS 속성선택자 → 편집기에서 값 바꾸면 즉시 반영, JS 미개입). 검은 영역이 이 폭까지 넓어지고 항목은 중앙에 모임 */
.mm-belief[data-maxw="900"]{max-width:900px;}
.mm-belief[data-maxw="1000"]{max-width:1000px;}
.mm-belief[data-maxw="1100"]{max-width:1100px;}
.mm-belief[data-maxw="1200"]{max-width:1200px;}
.mm-belief[data-maxw="full"]{max-width:none;}
.mm-belief__item{display:flex;flex-direction:column;align-items:center;gap:8px;}
.mm-belief__big{font-family:'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif;font-size:clamp(30px,9vw,52px);font-weight:900;line-height:1;letter-spacing:-.03em;color:#FAF9F7;font-variant-numeric:tabular-nums;}
.mm-belief[data-theme="라이트"] .mm-belief__big{color:#1A1A1A;}
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
```
### JS 탭
```js
(function(){
  function fmt(n){return n.toLocaleString('en-US');}
  function run(root){
    var bigs=root.querySelectorAll('.mm-belief__big');
    for(var i=0;i<bigs.length;i++)(function(el){
      var raw=(el.textContent||'').trim();
      var suf=el.getAttribute('data-suffix')||'';
      if(/^[\d,]+$/.test(raw)){ // 숫자 → 카운트업
        var target=parseInt(raw.replace(/,/g,''),10)||0,t0=null,dur=1300;
        function step(ts){if(!t0)t0=ts;var p=Math.min((ts-t0)/dur,1),e=1-Math.pow(1-p,3);el.textContent=fmt(Math.round(target*e))+suf;if(p<1)requestAnimationFrame(step);}
        requestAnimationFrame(step);
      } else { el.textContent=raw+suf; } // 글자 → 그대로(+단위)
    })(bigs[i]);
  }
  function initOne(root){
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
