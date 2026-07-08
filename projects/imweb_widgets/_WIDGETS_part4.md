# 🧩 MAMORU 아임웹 위젯 — Part 4/5 (n10_icon_nav ~ t07_spotlight)

> 각 위젯=HTML/CSS/JS 3탭. 🚫 삼중괄호 {{{ }}} · CSS탭 {{변수}} · 인라인 style="…{{}}…"(동적값은 data-*+JS) · 인라인 on*= 금지. 이 파일 10종.

## 📑 이 파일의 위젯
- N10. 아이콘 칩 내비 — `n10_icon_nav`
- N11. 유튜브 썸네일 배너 — `n11_youtube_banner`
- N12. 카테고리 이미지 그리드 (트렌디) — `n12_category_grid`
- T1. 스크롤 스토리텔링 (스티키) — `t01_storytelling`
- T2. 복원 인터랙티브 타임라인 — `t02_repair_timeline`
- T3. 가위 해부도 핫스팟 — `t03_hotspot`
- T4. 3D 틸트 제품 카드 — `t04_tilt_cards`
- T5. 무한 마퀴 띠 — `t05_marquee`
- T6. 라이트박스 사례 갤러리 — `t06_lightbox_gallery`
- T7. 마우스 스포트라이트 히어로 — `t07_spotlight`

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
{{!-- @name widgetInfo @type outlined-textfield @default "단계 카드 가로 진열. 단계명=주제 + 장면 설명. PC 한 줄, 모바일 가로 스크롤(우측 peek)로 세로 공간 절약." @label "ℹ️ 위젯 설명(참고용·수정 불필요)" --}}
<span style="display:none">{{widgetInfo}}</span>
<!-- ═══════════════════════════════════════════════════════════════
  📦 MAMORU 커스텀 위젯 — 단계 타임라인 (가로형)
  📍 아임웹 디자인모드 → 커스텀 위젯 → HTML 탭
  📝 수거→검수→수리→출고. 단계명=주제 + 장면 설명. PC 한 줄 / 모바일 가로 스크롤
  🚫 fetch·iframe 0
═══════════════════════════════════════════════════════════════ -->
{{!-- @name title @type outlined-textfield @default "복원 과정" @label "제목(선택)" --}}
{{!-- @name maxw @type outlined-textfield @default "" @label "가로 최대폭 — 숫자 자유 입력(예 1280 · 비우면 꽉 채움). 영역 확장해도 이 값에서 멈춤" --}}
{{!-- @name steps @type item @label "단계" --}}
<div class="mm-tl" data-maxw="{{maxw}}">
  <p class="mm-tl__title">{{title}}</p>
  <div class="mm-tl__scroll">
    <ol class="mm-tl__list">
    {{#each steps}}
      {{!-- @name stepName @type outlined-textfield @default "단계명" @label "단계명(주제)" --}}
      {{!-- @name stepImage @type image @label "장면 이미지(선택) — 권장 800×600px" --}}
      {{!-- @name stepDesc @type text-editor @default "<p>장면 설명</p>" @label "설명(장면)" --}}
      <li class="mm-tl__step">
        <span class="mm-tl__idx" aria-hidden="true"></span>
        <h3 class="mm-tl__name">{{stepName}}</h3>
        <img class="mm-tl__img" src="{{stepImage}}" alt="">
        <div class="mm-tl__desc">{{stepDesc}}</div>
      </li>
    {{/each}}
    </ol>
  </div>
</div>
```
### CSS 탭
```css
.mm-tl{box-sizing:border-box;--tl-gap:clamp(16px,3vw,28px);max-width:960px;margin:0 auto;padding:clamp(20px,4vw,32px) 0;font-family:'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif;color:#1A1A1A;}
.mm-tl__title{margin:0 0 clamp(16px,3vw,26px);padding:0 20px;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(20px,5vw,28px);font-weight:800;letter-spacing:-.02em;}
.mm-tl__title:empty{display:none;}
/* 가로 스크롤 컨테이너: PC 다 들어오면 균등, 모바일 넘치면 좌측부터 스크롤 + 우측 카드 peek */
.mm-tl__scroll{overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;scroll-snap-type:x proximity;padding:2px 20px;}
.mm-tl__scroll::-webkit-scrollbar{display:none;}
.mm-tl__list{counter-reset:tl;list-style:none;margin:0;padding:0;display:flex;flex-wrap:nowrap;gap:var(--tl-gap);}
/* 단계 카드: PC=균등(flex:1), 모바일=고정폭+가로스크롤 */
.mm-tl__step{position:relative;counter-increment:tl;flex:1 1 0;min-width:190px;scroll-snap-align:start;}
/* 인덱스 우측 커넥터 라인(다음 단계로 이어짐) */
.mm-tl__step:not(:last-child)::after{content:'';position:absolute;top:clamp(15px,3.4vw,20px);left:calc(clamp(30px,7vw,44px) + 14px);right:calc(-1 * var(--tl-gap) + 4px);height:1px;background:#D4D0CB;}
.mm-tl__idx{display:block;font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(30px,7vw,44px);font-weight:900;line-height:1;letter-spacing:-.04em;color:#1A1A1A;}
.mm-tl__idx::before{content:counter(tl,decimal-leading-zero);}
.mm-tl__name{margin:clamp(10px,2.2vw,16px) 0 8px;font-size:clamp(15px,2.4vw,19px);font-weight:800;letter-spacing:-.01em;}
/* 이미지 원본 비율 그대로(크롭X) — 높이는 이미지 비율에 맞춰 자동. 같은 비율 이미지 넣으면 카드 높이 정렬됨 */
.mm-tl__img{width:100%;height:auto;border-radius:10px;background:#F5F3F0;margin:0 0 10px;display:block;}
.mm-tl__img[src=""],.mm-tl__img:not([src]){display:none;}
.mm-tl__desc{font-size:clamp(13px,1.9vw,15px);line-height:1.6;color:#4A4A4A;}
.mm-tl__desc p{margin:0 0 .15em;}
/* 모바일: 카드 고정폭 → 가로 스크롤 + 우측 카드 살짝 보임(peek). 세로 공간 절약 */
@media (max-width:768px){
  .mm-tl__step{flex:0 0 76%;min-width:0;}
  .mm-tl__step:not(:last-child)::after{display:none;}
}
/* 등장(위젯이 화면에 들어오면 카드 순차 페이드) */
@media (prefers-reduced-motion:no-preference){
  .mm-tl__step{opacity:0;transform:translateY(16px);transition:opacity .5s cubic-bezier(.4,0,.2,1),transform .5s cubic-bezier(.4,0,.2,1);}
  .mm-tl__step.is-in{opacity:1;transform:none;}
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
  function initOne(root){
    applyMaxw(root);
    /* 최대폭 바뀌면 즉시 반영(편집기 실시간) */
    if('MutationObserver' in window){ new MutationObserver(function(){applyMaxw(root);}).observe(root,{attributes:true,attributeFilter:['data-maxw']}); }
    /* 위젯이 화면에 들어오면 카드 순차 페이드(가로형이라 컨테이너 기준) */
    var steps=root.querySelectorAll('.mm-tl__step');
    function reveal(){ for(var k=0;k<steps.length;k++){ steps[k].style.transitionDelay=(k*0.08)+'s'; steps[k].classList.add('is-in'); } }
    if('IntersectionObserver' in window){
      var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){reveal();io.disconnect();}});},{threshold:.15});
      io.observe(root);
    }else{ reveal(); }
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
