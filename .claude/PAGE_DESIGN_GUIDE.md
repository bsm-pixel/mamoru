# MAMORU 페이지 제작 가이드

> 모든 페이지(아임웹 코드위젯, 독립 페이지, 간편리뷰 등)에 일관된 룩앤필을 보장하기 위한 실행 규격.
> 색상·타이포그래피 상세는 `BRAND_COLOR_SYSTEM.md`를 참조한다.

---

## 1. 개요

### 적용 범위
- 아임웹 코드위젯 (iframe 삽입형)
- GitHub Pages 독립 페이지
- 향후 추가되는 모든 MAMORU 웹 페이지

### 관련 문서
| 문서 | 역할 |
|------|------|
| `BRAND_COLOR_SYSTEM.md` | 색상 팔레트, 타이포 스케일, 금지 조합 |
| `ADDENDUM_IMWEB.md` | 아임웹 제약사항, CSS/JS 충돌 방지 |
| 이 문서 | 레이아웃, 컴포넌트, 반응형, iframe 통신 |

---

## 2. 페이지 유형 분류

### A. 쇼윈도형 — max-width: 1200px
- 이미지·그리드·카드 중심, 시각 밀도 높음
- 레퍼런스: `projects/main/page_main.html`
- 특징: 모바일 다크 중심 → PC 라이트 전환

### B. 리딩형 — max-width: 960px
- 텍스트·카드·프로세스 중심, 독서 흐름 최적화
- 레퍼런스: `projects/consulting/page_guide.html`, `projects/reviews/page_reviews.html`
- 특징: 히어로만 다크, 나머지 크림

---

## 3. 브레이크포인트 체계

| 단계 | 범위 | 역할 | 미디어쿼리 |
|------|------|------|-----------|
| 소형폰 | ≤360px | 절대 최소 보장 | `@media (max-width: 360px)` |
| 모바일 | 361–767px | base (모바일 우선) | 기본 스타일 |
| 태블릿 | 768–1023px | 2컬럼, 패딩 확대 | `@media (min-width: 768px)` |
| PC | 1024–1439px | 3컬럼, 전폭 히어로 | `@media (min-width: 1024px)` |
| 와이드 | 1440px+ | 최대 타이포·패딩 | `@media (min-width: 1440px)` |
| 호버 분리 | — | PC 호버 효과 격리 | `@media (hover: hover)` |

---

## 4. PC 레이아웃 규칙

### 콘텐츠 너비 단계
| 역할 | 너비 | 사용처 |
|------|------|--------|
| 전폭 | 100% | 히어로 다크 배경 |
| 쇼윈도 | 1200px | 메인페이지 콘텐츠 |
| 리딩 | 960px | 상담안내·후기·간편리뷰 |
| 집중 | 640px | 프로세스, 원칙 리스트, CTA 영역 |

### 수평 패딩 단계
| 브레이크포인트 | 패딩 |
|--------------|------|
| 모바일 (base) | `20px` |
| 태블릿 (768px+) | `32px` |
| PC (1024px+) | `48px` |
| 와이드 (1440px+) | `64px` |

### 수직 패딩 참고
- 히어로: 모바일 `60px 20px` → PC `72px 48px`
- 일반 섹션: 모바일 `48px 20px` → PC `64px 48px`
- 섹션 간 간격은 8px 단위(16/24/32/40/48)

---

## 5. 모바일 레이아웃 규칙

- **1컬럼 기본**, 수평 패딩 `20px`
- 터치 타깃 최소 **48px** (패딩 포함 높이)
- active 피드백: `scale(0.96)` ~ `scale(0.97)`, `0.15s ease`
- 본문 최소 `14px`, 라벨 최소 `11px` (360px 이하에서)
- `-webkit-tap-highlight-color: transparent` 필수
- 수평 스크롤 카드 리스트: `overflow-x: auto`, 스크롤바 숨김

### 스크롤바 숨김
```css
.scroll-container {
  overflow-x: auto;
  scrollbar-width: none;
  -ms-overflow-style: none;
}
.scroll-container::-webkit-scrollbar { display: none; }
```

---

## 6. 다크/라이트 영역 전략

### 일반 페이지 (리딩형)
```
히어로: 다크(#181725) — 텍스트 #F2F2EA
본문: 크림(#F2F2EA) — 텍스트 #181725
카드: #FAFAF5 — 텍스트 #181725
```
- 60/30/10 비율 준수
- 히어로만 다크, 나머지 모두 크림 계열

### 메인페이지 (쇼윈도형 — 예외)
```
모바일: 다크 중심 → 일부 라이트 섹션(상품/가이드)
PC:     페이지 전체 라이트(#F2F2EA) + 다크 카드 컨테이너
```
- PC에서 다크 섹션을 카드 형태로 라이트 배경 위에 배치

### 배경색 참조
| 역할 | HEX | CSS 변수 |
|------|-----|---------|
| 페이지 배경 | `#F2F2EA` | `--mm-bg-light` |
| 카드 배경 | `#FAFAF5` | `--mm-bg-card` |
| 미묘한 구분 | `#EDEEE6` | (인라인) |
| 다크 배경 | `#181725` | `--mm-bg-dark` |
| 다크 딥 | `#100F1C` | `--mm-bg-darker` |
| 다크 카드 | `#1F1E2D` ~ `#2A2940` | `--mm-bg-card-dark` |

---

## 7. 컴포넌트 표준

### 7-1. 카드
```css
.mm-card {
  background: var(--mm-bg-card);        /* #FAFAF5 */
  border: 1px solid var(--mm-border-l); /* rgba(24,23,37,0.08) */
  border-radius: var(--mm-radius-lg);   /* 8px */
  padding: 20px;                        /* 모바일 */
  transition: all var(--mm-ease);       /* 0.25s ease */
}
/* 태블릿+ */
@media (min-width: 768px) {
  .mm-card { padding: 24px; }
}
/* 호버 */
@media (hover: hover) {
  .mm-card:hover {
    box-shadow: 0 4px 16px rgba(24,23,37,0.06);
    border-color: rgba(212,97,62,0.15);
  }
}
```
- 카드 내 이미지: `aspect-ratio: 16/9`, `border-radius: 6px`, `object-fit: cover`

### 7-2. 탭/필터 (Pill)
```css
/* 컨테이너 */
.mm-tabs {
  display: flex;
  justify-content: center;
  gap: 8px;
}
/* 개별 탭 */
.mm-tab {
  padding: 8px 20px;
  border-radius: 999px;
  border: 1px solid var(--mm-border-l);
  background: transparent;
  font-size: 13px;
  font-weight: 500;
  color: var(--mm-text-d2);
  transition: all var(--mm-ease-fast);
  cursor: pointer;
}
.mm-tab--active {
  background: var(--mm-gold);
  border-color: var(--mm-gold);
  color: var(--mm-text-w);
  font-weight: 600;
}
.mm-tab:active { transform: scale(0.96); }

@media (min-width: 768px) {
  .mm-tabs { gap: 10px; }
  .mm-tab { padding: 9px 24px; font-size: 14px; }
}
@media (min-width: 1024px) {
  .mm-tab { padding: 10px 28px; font-size: 15px; }
}
@media (hover: hover) {
  .mm-tab:not(.mm-tab--active):hover {
    border-color: var(--mm-gold);
    color: var(--mm-gold);
  }
}
```
- 모바일에서 탭이 많을 경우: `overflow-x: auto` + 스크롤바 숨김
- 다크 배경 위 탭: `border: none`, 비활성 `color: rgba(242,242,234,0.52)`

### 7-3. CTA 버튼
```css
/* Primary (테라코타) */
.mm-cta--primary {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 14px 28px;
  min-height: 52px;
  background: var(--mm-gold);
  border: none;
  border-radius: 999px;
  font-size: 15px;
  font-weight: 700;
  color: #fff;
  box-shadow: 0 4px 20px rgba(212,97,62,0.3);
  transition: transform var(--mm-ease-fast), box-shadow var(--mm-ease-fast);
  cursor: pointer;
}
.mm-cta--primary:active { transform: scale(0.97); }

/* Secondary (아웃라인) */
.mm-cta--secondary {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 14px 28px;
  background: transparent;
  border: 1px solid var(--mm-border-l);
  border-radius: 999px;
  font-size: 15px;
  font-weight: 500;
  color: var(--mm-text-d2);
  transition: all var(--mm-ease-fast);
  cursor: pointer;
}
.mm-cta--secondary:active { transform: scale(0.97); }

@media (min-width: 768px) {
  .mm-cta--primary,
  .mm-cta--secondary {
    padding: 16px 32px;
    min-height: 56px;
    font-size: 17px;
  }
}
@media (hover: hover) {
  .mm-cta--primary:hover {
    background: var(--mm-gold-dark);
    transform: translateY(-2px);
    box-shadow: 0 8px 24px rgba(212,97,62,0.25);
  }
  .mm-cta--secondary:hover {
    border-color: var(--mm-gold);
    color: var(--mm-gold);
  }
}
```

### 7-4. 섹션 헤더
```css
.mm-section-head {
  display: flex;
  flex-direction: column;
  align-items: center;
  text-align: center;
  margin-bottom: 32px;
}
/* 테라코타 바 */
.mm-section-head__bar {
  width: 28px;
  height: 3px;
  border-radius: 2px;
  background: var(--mm-gold);
  margin-bottom: 16px;
  box-shadow: 0 1px 6px rgba(212,97,62,0.2);
}
/* 타이틀 */
.mm-section-head__title {
  font-size: 20px;
  font-weight: 700;
  letter-spacing: -0.02em;
  line-height: 1.25;
  color: var(--mm-text-d);
  margin-bottom: 8px;
}
/* 설명 */
.mm-section-head__desc {
  font-size: 14px;
  font-weight: 400;
  color: var(--mm-text-d2);
  letter-spacing: -0.01em;
  line-height: 1.55;
}

@media (min-width: 768px) {
  .mm-section-head { margin-bottom: 36px; }
  .mm-section-head__title { font-size: 26px; }
  .mm-section-head__desc { font-size: 15px; }
}
@media (min-width: 1024px) {
  .mm-section-head { margin-bottom: 44px; }
  .mm-section-head__bar { width: 32px; }
  .mm-section-head__title { font-size: 30px; letter-spacing: -0.025em; }
  .mm-section-head__desc { font-size: 16px; }
}
```

### 7-5. 그리드
```css
.mm-grid {
  display: grid;
  grid-template-columns: 1fr;
  gap: 16px;
}
@media (min-width: 768px) {
  .mm-grid {
    grid-template-columns: repeat(2, 1fr);
    gap: 20px;
  }
}
@media (min-width: 1024px) {
  .mm-grid {
    grid-template-columns: repeat(3, 1fr);
    gap: 24px;
  }
}
```
- 쇼윈도형에서 4컬럼 필요 시: `repeat(4, 1fr)` (1024px+)

---

## 8. 호버/인터랙션 규칙

### 전환 속도
| 변수 | 값 | 용도 |
|------|-----|------|
| `--mm-ease-fast` | `0.15s ease` | 버튼, 탭, 즉각 피드백 |
| `--mm-ease` | `0.25s ease` | 카드, 일반 전환 |
| `--mm-ease-slow` | `0.4s ease` | 이미지 확대, 느린 노출 |

### 호버 효과 (`@media (hover: hover)` 내에서만)
| 요소 | 효과 | 속도 |
|------|------|------|
| 카드 | `box-shadow: 0 4px 16px rgba(24,23,37,0.06)` | 0.25s |
| CTA 버튼 | `translateY(-2px)` + shadow 증가 | 0.15s |
| 탭 (비활성) | `border-color` + `color` → 테라코타 | 0.15s |
| 이미지 카드 | `scale(1.04)` (이미지만) | 0.4s |
| 아이콘 | `scale(1.08)` | 0.25s |
| 링크/화살표 | `translateX(3px)` | 0.15s |

### 터치 피드백 (모든 디바이스)
| 요소 | `:active` 효과 |
|------|--------------|
| CTA 버튼 | `scale(0.97)` |
| 탭 | `scale(0.96)` |
| 카드 | `scale(0.97)` ~ `scale(0.98)` |
| 소형 버튼 | `scale(0.97)` |

---

## 9. iframe 통신 표준

### 메시지 타입
| 타입 | 방향 | 페이로드 | 용도 |
|------|------|---------|------|
| `MAMORU_IFRAME_SIZE` | 자식→부모 | `{ height: number }` | 높이 동기화 |
| `MAMORU_NAVIGATE` | 자식→부모 | `{ url: string }` | 네비게이션 위임 |
| `MAMORU_SCROLL_TOP` | 자식→부모 | (없음) | 스크롤 최상단 |
| `REQUEST_HEIGHT` | 부모→자식 | (없음) | 높이 재요청 |

### 표준 구현
```javascript
(function() {
  var lastH = 0;

  function getHeight() {
    // 특정 래퍼가 있으면 그 높이를 우선 사용
    var wrap = document.querySelector('.mm-page');
    if (wrap) {
      var r = wrap.getBoundingClientRect();
      return Math.ceil(r.height + 100);
    }
    return Math.ceil(document.documentElement.scrollHeight + 100);
  }

  function send() {
    var h = getHeight();
    if (h > 0 && Math.abs(h - lastH) > 5) {
      lastH = h;
      window.parent.postMessage({
        type: 'MAMORU_IFRAME_SIZE',
        height: h
      }, '*');
    }
  }

  // 높이 감지: ResizeObserver 우선
  if (typeof ResizeObserver !== 'undefined') {
    new ResizeObserver(send).observe(document.body);
  } else {
    setInterval(send, 800);
  }

  // 초기 전송 (지연 포함)
  send();
  window.addEventListener('load', send);
  setTimeout(send, 500);
  setTimeout(send, 1000);

  // 부모 요청 대응
  window.addEventListener('message', function(e) {
    if (e.data && e.data.type === 'REQUEST_HEIGHT') send();
  });

  // 네비게이션 위임
  document.addEventListener('click', function(e) {
    var a = e.target.closest('a[href]');
    if (!a) return;
    var href = a.getAttribute('href');
    if (!href || href.startsWith('http') || href.startsWith('mailto:') || href.startsWith('tel:')) return;
    e.preventDefault();
    window.parent.postMessage({ type: 'MAMORU_NAVIGATE', url: href }, '*');
  });
})();
```

### 독립 실행 조건
```javascript
// iframe 내부인 경우에만 통신 활성화
if (window.parent !== window) {
  initIframeComm();
}
```

---

## 10. 보일러플레이트 템플릿

아래 HTML을 복사하여 새 페이지 시작점으로 사용한다. `[리딩형]` 기준이며, 쇼윈도형은 `--mm-max-w`를 `1200px`로 변경한다.

```html
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MAMORU — [페이지명]</title>

<!-- Pretendard 폰트 (preload → stylesheet 전환) -->
<link rel="preconnect" href="https://cdn.jsdelivr.net" crossorigin>
<link rel="preload"
  href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css"
  as="style"
  onload="this.onload=null;this.rel='stylesheet'">
<noscript>
  <link href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css"
    rel="stylesheet">
</noscript>

<style>
/* ── Reset ── */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
img{display:block;max-width:100%}
a{text-decoration:none;color:inherit}
button{font:inherit;cursor:pointer;border:none;background:none}

/* ── CSS 변수 (Terracotta Editorial) ── */
:root{
  /* Primary */
  --mm-bg-dark:#181725;
  --mm-bg-darker:#100F1C;
  --mm-bg-card-dark:#2A2940;
  --mm-bg-light:#F2F2EA;
  --mm-bg-warm:#F2F2EA;
  --mm-bg-card:#FAFAF5;

  --mm-gold:#D4613E;
  --mm-gold-dark:#B85232;
  --mm-gold-soft:rgba(212,97,62,0.14);
  --mm-gold-glow:rgba(212,97,62,0.07);

  /* Text on Dark */
  --mm-text-w:#F2F2EA;
  --mm-text-w2:rgba(242,242,234,0.6);
  --mm-text-w3:rgba(242,242,234,0.52);
  --mm-text-w4:rgba(242,242,234,0.45);

  /* Text on Light */
  --mm-text-d:#181725;
  --mm-text-d2:#6B6980;
  --mm-text-d3:rgba(24,23,37,0.55);

  /* Borders */
  --mm-border-d:rgba(255,255,255,0.08);
  --mm-border-l:rgba(24,23,37,0.08);

  /* Layout */
  --mm-max-w:960px;    /* 리딩형 960px / 쇼윈도형은 1200px */
  --mm-radius:6px;
  --mm-radius-lg:8px;

  /* Transitions */
  --mm-ease-fast:0.15s ease;
  --mm-ease:0.25s ease;
  --mm-ease-slow:0.4s ease;
}

body{
  font-family:'Pretendard Variable',-apple-system,BlinkMacSystemFont,system-ui,sans-serif;
  background:var(--mm-bg-light);
  color:var(--mm-text-d);
  line-height:1.6;
  -webkit-font-smoothing:antialiased;
  -webkit-tap-highlight-color:transparent;
}

/* ── 레이아웃 래퍼 ── */
.mm-page{
  max-width:100%;
  overflow-x:hidden;
}
.mm-inner{
  max-width:var(--mm-max-w);
  margin:0 auto;
  padding:0 20px;
}
@media(min-width:768px){.mm-inner{padding:0 32px}}
@media(min-width:1024px){.mm-inner{padding:0 48px}}
@media(min-width:1440px){.mm-inner{padding:0 64px}}

/* ── 히어로 (다크) ── */
.mm-hero{
  background:var(--mm-bg-dark);
  padding:60px 20px;
  text-align:center;
}
.mm-hero__bar{
  width:28px;height:3px;
  border-radius:2px;
  background:var(--mm-gold);
  margin:0 auto 16px;
}
.mm-hero__title{
  font-size:28px;font-weight:800;
  line-height:1.3;letter-spacing:-0.03em;
  color:var(--mm-text-w);
  margin-bottom:16px;
}
.mm-hero__desc{
  font-size:14px;font-weight:400;
  color:var(--mm-text-w2);
  line-height:1.6;
}
@media(min-width:768px){
  .mm-hero{padding:72px 48px}
  .mm-hero__title{font-size:42px;line-height:1.2;letter-spacing:-0.04em}
  .mm-hero__desc{font-size:16px}
}
@media(max-width:360px){
  .mm-hero__title{font-size:24px}
}

/* ── 섹션 ── */
.mm-section{
  padding:48px 0;
}
.mm-section-head{
  display:flex;flex-direction:column;
  align-items:center;text-align:center;
  margin-bottom:32px;
}
.mm-section-head__bar{
  width:28px;height:3px;
  border-radius:2px;
  background:var(--mm-gold);
  margin-bottom:16px;
  box-shadow:0 1px 6px rgba(212,97,62,0.2);
}
.mm-section-head__title{
  font-size:20px;font-weight:700;
  letter-spacing:-0.02em;line-height:1.25;
  color:var(--mm-text-d);margin-bottom:8px;
}
.mm-section-head__desc{
  font-size:14px;font-weight:400;
  color:var(--mm-text-d2);
  letter-spacing:-0.01em;line-height:1.55;
}
@media(min-width:768px){
  .mm-section{padding:56px 0}
  .mm-section-head{margin-bottom:36px}
  .mm-section-head__title{font-size:26px}
  .mm-section-head__desc{font-size:15px}
}
@media(min-width:1024px){
  .mm-section{padding:64px 0}
  .mm-section-head{margin-bottom:44px}
  .mm-section-head__bar{width:32px}
  .mm-section-head__title{font-size:30px;letter-spacing:-0.025em}
  .mm-section-head__desc{font-size:16px}
}

/* ── 카드 ── */
.mm-card{
  background:var(--mm-bg-card);
  border:1px solid var(--mm-border-l);
  border-radius:var(--mm-radius-lg);
  padding:20px;
  transition:all var(--mm-ease);
}
@media(min-width:768px){.mm-card{padding:24px}}
@media(hover:hover){
  .mm-card:hover{
    box-shadow:0 4px 16px rgba(24,23,37,0.06);
    border-color:rgba(212,97,62,0.15);
  }
}

/* ── 그리드 ── */
.mm-grid{
  display:grid;
  grid-template-columns:1fr;
  gap:16px;
}
@media(min-width:768px){
  .mm-grid{grid-template-columns:repeat(2,1fr);gap:20px}
}
@media(min-width:1024px){
  .mm-grid{grid-template-columns:repeat(3,1fr);gap:24px}
}

/* ── 탭 (Pill) ── */
.mm-tabs{
  display:flex;justify-content:center;
  gap:8px;flex-wrap:wrap;
}
.mm-tab{
  padding:8px 20px;
  border-radius:999px;
  border:1px solid var(--mm-border-l);
  background:transparent;
  font-size:13px;font-weight:500;
  color:var(--mm-text-d2);
  transition:all var(--mm-ease-fast);
}
.mm-tab--active{
  background:var(--mm-gold);
  border-color:var(--mm-gold);
  color:var(--mm-text-w);
  font-weight:600;
}
.mm-tab:active{transform:scale(0.96)}
@media(min-width:768px){.mm-tabs{gap:10px}.mm-tab{padding:9px 24px;font-size:14px}}
@media(min-width:1024px){.mm-tab{padding:10px 28px;font-size:15px}}
@media(hover:hover){
  .mm-tab:not(.mm-tab--active):hover{
    border-color:var(--mm-gold);color:var(--mm-gold);
  }
}

/* ── CTA 버튼 ── */
.mm-cta{
  display:inline-flex;align-items:center;justify-content:center;
  border-radius:999px;
  font-weight:700;
  transition:all var(--mm-ease-fast);
}
.mm-cta--primary{
  padding:14px 28px;min-height:52px;
  background:var(--mm-gold);border:none;
  color:#fff;font-size:15px;
  box-shadow:0 4px 20px rgba(212,97,62,0.3);
}
.mm-cta--secondary{
  padding:14px 28px;
  background:transparent;
  border:1px solid var(--mm-border-l);
  color:var(--mm-text-d2);font-size:15px;
  font-weight:500;
}
.mm-cta:active{transform:scale(0.97)}
@media(min-width:768px){
  .mm-cta--primary,.mm-cta--secondary{
    padding:16px 32px;min-height:56px;font-size:17px;
  }
}
@media(hover:hover){
  .mm-cta--primary:hover{
    background:var(--mm-gold-dark);
    transform:translateY(-2px);
    box-shadow:0 8px 24px rgba(212,97,62,0.25);
  }
  .mm-cta--secondary:hover{
    border-color:var(--mm-gold);color:var(--mm-gold);
  }
}

/* ── CTA 고정 영역 ── */
.mm-cta-area{
  display:flex;gap:8px;
  max-width:640px;margin:40px auto 0;
}
.mm-cta-area .mm-cta{flex:1}
</style>
</head>

<body>
<div class="mm-page">

  <!-- ▸ 히어로 (다크) -->
  <section class="mm-hero">
    <div class="mm-hero__bar"></div>
    <h1 class="mm-hero__title">[히어로 타이틀]</h1>
    <p class="mm-hero__desc">[히어로 설명]</p>
  </section>

  <!-- ▸ 콘텐츠 섹션 예시 -->
  <section class="mm-section">
    <div class="mm-inner">
      <div class="mm-section-head">
        <span class="mm-section-head__bar"></span>
        <h2 class="mm-section-head__title">[섹션 타이틀]</h2>
        <p class="mm-section-head__desc">[섹션 설명]</p>
      </div>

      <!-- 탭 필터 (필요 시) -->
      <div class="mm-tabs" style="margin-bottom:24px">
        <button class="mm-tab mm-tab--active">전체</button>
        <button class="mm-tab">카테고리 A</button>
        <button class="mm-tab">카테고리 B</button>
      </div>

      <!-- 그리드 카드 -->
      <div class="mm-grid">
        <div class="mm-card">[카드 내용]</div>
        <div class="mm-card">[카드 내용]</div>
        <div class="mm-card">[카드 내용]</div>
      </div>
    </div>
  </section>

  <!-- ▸ CTA 영역 -->
  <section class="mm-section">
    <div class="mm-inner">
      <div class="mm-cta-area">
        <button class="mm-cta mm-cta--primary">[주요 액션]</button>
        <button class="mm-cta mm-cta--secondary">[보조 액션]</button>
      </div>
    </div>
  </section>

</div>

<!-- ── iframe 통신 ── -->
<script>
(function(){
  var lastH=0;
  function getH(){
    var w=document.querySelector('.mm-page');
    if(w){var r=w.getBoundingClientRect();return Math.ceil(r.height+100)}
    return Math.ceil(document.documentElement.scrollHeight+100);
  }
  function send(){
    var h=getH();
    if(h>0&&Math.abs(h-lastH)>5){
      lastH=h;
      window.parent.postMessage({type:'MAMORU_IFRAME_SIZE',height:h},'*');
    }
  }
  if(window.parent===window)return; // 독립 실행 시 skip
  if(typeof ResizeObserver!=='undefined'){
    new ResizeObserver(send).observe(document.body);
  }else{setInterval(send,800)}
  send();
  window.addEventListener('load',send);
  setTimeout(send,500);setTimeout(send,1000);
  window.addEventListener('message',function(e){
    if(e.data&&e.data.type==='REQUEST_HEIGHT')send();
  });
  document.addEventListener('click',function(e){
    var a=e.target.closest('a[href]');
    if(!a)return;
    var href=a.getAttribute('href');
    if(!href||href.startsWith('http')||href.startsWith('mailto:')||href.startsWith('tel:'))return;
    e.preventDefault();
    window.parent.postMessage({type:'MAMORU_NAVIGATE',url:href},'*');
  });
})();
</script>
</body>
</html>
```

---

## 11. 체크리스트

### 제작 완료 후 검증 항목

**레이아웃**
- [ ] 모바일 375px에서 콘텐츠 잘림/넘침 없음
- [ ] PC 1440px에서 콘텐츠가 max-width 내에 수렴
- [ ] 수평 패딩 단계 준수 (20→32→48→64px)
- [ ] 히어로 전폭 + 콘텐츠 max-width 분리 확인

**타이포그래피**
- [ ] 360px 이하에서 본문 14px, 라벨 11px 이상
- [ ] 섹션 타이틀 20→26→30px 스케일 적용
- [ ] Pretendard Variable 로드 확인 (개발자도구 > Computed)

**컬러**
- [ ] 순수 `#000` / `#FFF` 사용 없음
- [ ] 60/30/10 비율 시각 확인
- [ ] 크림 배경 위 테라코타 텍스트 사용 없음 (대비 부족)

**인터랙션**
- [ ] 모바일: 터치 active `scale()` 피드백 동작
- [ ] PC: 호버 효과가 `@media (hover: hover)` 내에 격리
- [ ] CTA 버튼 min-height 52px(모바일) / 56px(태블릿+)
- [ ] `focus-visible` 아웃라인 존재

**iframe**
- [ ] `MAMORU_IFRAME_SIZE` 메시지 발신 확인
- [ ] ResizeObserver 또는 setInterval fallback 동작
- [ ] 내부 링크 클릭 시 `MAMORU_NAVIGATE` 발신 확인
- [ ] 독립 실행(비-iframe) 시 에러 없음

**성능**
- [ ] CSS 네임스페이스 `.mm-` 접두어 사용
- [ ] 전역 JS 변수 없음 (IIFE 또는 네임스페이스)
- [ ] 폰트 preload 적용

---

## 변경 이력

| 날짜 | 변경 내용 |
|------|----------|
| 2026-02-14 | 초안 작성: 3개 페이지(메인·상담안내·후기) 패턴 종합 |
