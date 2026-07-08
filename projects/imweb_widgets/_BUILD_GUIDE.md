# 🛠️ 아임웹 커스텀 위젯 제작 가이드 (필수 규칙 SSOT)

> 가이드 정밀 답변(2026-06-30) + 우리 실측을 종합한 **제작 시 필수 준수 규칙**.
> 표기: ✅확정(FAQ/실측) · ⚠️주의(미검증/불일치) · 🔧우리 코드 반영사항.

---

## 🚨🚨 위젯 제작 필독 — 사장님 확정 방향 (모든 위젯 이걸로. 2026-07-08)

**새 위젯 만들거나 고칠 때 이 16개를 먼저 훑고 시작한다.**

### A. 아임웹 저장거부 회피 (안 지키면 "유효하지 않은 코드")
1. **CSS값 자리 `{{}}` 전면 금지** — CSS 탭·인라인 `style="…{{}}…"` **둘 다 거부**. 동적 스타일값은 **`data-*` 속성**으로만.
2. **삼중괄호 `{{{}}}` 금지**, 인라인 `on*=` 금지, `<script>/<style>/<iframe>/<form>` HTML탭 금지, `{{@index}}/{{../x}}/{{this.x}}` 금지.
3. **CSS `url(http/https/data:)`·`@import` 금지**(폰트=SEO 바디코드, 아이콘=인라인 SVG). JS `fetch/XHR/WebSocket/localStorage/cookie/location변경/eval` 금지.

### B. 패널값 → 스타일 적용 방식 (핵심)
4. **discrete(테마·정렬·각지게·방향 등)** = `data-*="{{v}}"` + **CSS 속성 선택자**(예 `[data-theme="라이트"]`). 편집기 즉시반영·깜빡임0.
5. **연속/자유숫자(가로폭·높이·폰트·열수·색)** = `data-*` + **JS + `MutationObserver`**(속성변경 감지→즉시 재적용). ⚠️아임웹 편집기는 패널값 바꿔도 **JS 자동 재실행 안 함** → Observer 필수.
6. 필드 **유형 바꾸려면 변수명을 새로**(예 `icon`→`iconImage`). 필드 구조 바뀌면 **위젯 삭제 후 재추가**(패널 고착).

### C. 레이아웃 (사장님 필수 요구)
7. **가로 최대폭** = 자유 숫자입력(예 1280) + `box-sizing:border-box` + `margin auto` + `applyMaxw`(비움/full=꽉참). **영역 100% 확장해도 설정폭에서 멈추고 중앙.**
8. **PC/모바일 각각** = 이미지·높이·폰트·열수를 뷰포트별(@media / CSS var)로.
9. **세로로 길어지는 목록형은 모바일 가로 스크롤 + 우측 카드 peek**(트랙/스크롤 컨테이너, 카드 flex 고정폭). PC=중앙 균등. **화면 세로 안 잡아먹게.**
10. **모바일 좌우 여백**(섹션 100% 확장 대응): **밴드형=좌우 16px**(배경은 border-box라 화면 끝까지 블리드, 콘텐츠만 인셋) · **가로스크롤형=좌측만**(우측 peek 유지). 흐름배치 이미지 블리드 위젯(d01·14)은 루트 패딩 제외(이너 패딩).
11. **히어로/배너 높이 = 이미지 비율 자동**: 비우면 **이미지 비율대로 높이 자동(PC=PC이미지 / 모바일=모바일이미지)**, 값 넣으면 고정+크롭. **이미지 있으면 min-height 바닥 제거**(하단 검정여백 방지). 낮은 배너는 **높이 비례 폰트 스케일**(`--cut-s`류)로 안 눌리게.

### D. 브랜드·디테일
12. **모노크롬 9색**(Void #1A1A1A ~ Cream #FAF9F7), Outfit/Pretendard, **채도 컬러 금지**. **모서리 각지게 기본**(라운드 안 함, 필요시만).
13. **빈 요소 숨김** `:empty`, **조건부 요소**(예 키커 있을 때만 가로선 `:empty + .blade`). 큰 한글 글씨 **획 보정**(is-ko).
14. 이미지 필드는 아임웹 **SVG 업로드 막히면 PNG/제외**. 단색 아이콘 테마색=**mask + background:currentColor**.

### E. 작업 마무리 (매번)
15. `_main_page` 사본 동기화 → 문서 재생성(`/tmp/regen.py`·`regen_all.py`) + **프리뷰 재생성 `py _gen_preview.py`** → 커밋·푸시.
16. **프리뷰 `projects/_design_lab/widgets_preview.html`** 로 PC/모바일·값입력 즉시 확인(아임웹 등록 전).

---

## 1. 🚫 절대 금지 (= 저장 시 "유효하지 않은 코드" 차단)

### HTML 탭
- ✅ **인라인 이벤트 핸들러 `on*=` 전면 금지** (`onerror` `onclick` `onload` `onmouseover` …) → **JS에서 `addEventListener`로.**
- ✅ **태그 금지**: `<script>` `<style>` `<iframe>` `<object>` `<embed>` `<form>` 및 `<!DOCTYPE>` `<html>` `<head>` `<body>` `<base>` `<meta>` `<link>`. (3탭 분리이므로 style/script는 각 탭에 — HTML 탭엔 마크업만)
- ✅ `href`/`src`에 `javascript:` · `data:` 시작 금지.

### JS 탭
- ✅ **금지**: `eval` · `new Function` · `fetch` · `XMLHttpRequest` · `WebSocket` · `localStorage`/`sessionStorage`/`indexedDB` · `document.cookie` · `window.location`/`history`/`navigator` 변경 · 외부 라이브러리/CDN.
- ⚠️ **`document.addEventListener`** — 프롬프트코어는 "금지, element/window에 바인딩" 주장. **그러나 우리 실측: `document.addEventListener('DOMContentLoaded', init)`로 d01 저장됨(허용).** → DOMContentLoaded는 OK. 단 안전하게 가려면 `window.addEventListener`도 가능.
- ⚠️ **정규식** — 금지 목록엔 없으나 명시도 없음. 복잡한 정규식 ReDoS 차단 가능성 → **되도록 문자열 파싱(indexOf/split)으로**(우리 d01 이미 적용).

### CSS 탭
- ✅ **외부 URL 전면 금지**: `url('https://...')`, `@import url(...)`, `url('data:...')`. → 폰트는 SEO>body code에 링크 후 `font-family`만 선언. 아이콘은 **인라인 SVG**(✅허용·권장).

---

## 2. ✅ 허용·권장 패턴 (이대로 쓰면 안전)
- `setTimeout`/`setInterval`(첫 인자 **함수만**, 문자열 금지) · `requestAnimationFrame` · `MutationObserver`(컨테이너 내부만, body+subtree 금지) · `IntersectionObserver` · `ResizeObserver`.
- `element.addEventListener` / `window.addEventListener('resize', fn)` · `window.innerWidth/innerHeight` · `window.getComputedStyle`.
- `document.createElement`(위험태그 제외) · `element.innerHTML`(인라인 SVG 주입 OK) · `querySelector/All` · `get/setAttribute`.
- `Math` `JSON` `String` `Array` 등 표준 내장.
- **격리**: CSS는 **Shadow DOM 격리**(전역과 안 섞임), JS `document`는 위젯 컨테이너로 스코핑된 프록시 → `document.querySelectorAll('.mm-x')`가 그 위젯 범위만 잡음(우리 패턴 OK). **위젯 간 통신 불가.**

---

## 3. 🎛️ 패널 @type 토큰 (★실측 우선)

| UI 종류 | 우리가 쓸 @type | 근거 |
|---|---|---|
| 텍스트 한 줄 | **`outlined-textfield`** | ✅실측 동작 |
| 줄글(서식) | `text-editor` | FAQ |
| 이미지 업로드 | **`image`** | ✅실측 (image-uploader 아님) |
| 색상+투명도 | **`color-picker`** | ✅실측 동작 |
| 드롭다운 | `select` | FAQ(미검증) |
| ON/OFF | `switch` | FAQ — ⚠️**반환값 `"true"/"false"` 문자열** → `data-*` + `=== 'true'` 비교(우리 적용) |
| 날짜/시간 | `date-picker` / `time-picker` | FAQ(미검증) ⚠️반환 형식 불명 → 유연 파싱 |
| 옵션버튼(탭) | ⚠️`segmented-control`=텍스트로 떨어짐(실측). 프롬프트코어는 **`segment`**가 정답이라 함 → **테스트 후 채택**. 현재는 **텍스트칸+"입력: 값" 라벨**로 우회 중 | 실측+코어 |
| 반복 | `item` | FAQ |

**🔑 패널 "유형"은 UI에서 바꾸는 게 가장 확실(2026-06-30 실측)**: 코드 붙이면 아임웹이 변수마다 패널 항목을 만들되 **유형을 "입력창"으로 기본 설정**(annotation `@type`은 image/color-picker/outlined-textfield 등 일부만 자동반영, **`segment` 등은 입력창으로 떨어짐**). → **버튼/토글/드롭다운이 필요하면 그 항목 클릭 → "항목 속성 → 유형" 드롭다운에서 옵션버튼·스위치·드롭다운 선택 → 값 입력 → 적용.** 코드 토큰 싸움 불필요. **우리 JS는 `data-*` 값을 어떤 유형이 와도(텍스트/스위치 `"true"`/옵션버튼 값) 받게 작성**한다.

**색상 @default = `#RRGGBBAA` 8자리 hex 필수** (예: 검정50% = `#1A1A1A80`). `rgba()`·3/6자리는 파싱 실패 시 `#000000FF`(불투명 검정)로 fallback → 🔧우리 color 기본값 8-hex로 교체.

**annotation 지원 7태그만**: `@name @type @default @label @placeholder @values @valueNames`. (정적 안내문·구분선·`@description` **없음** → 설명은 "안내용 textfield"로 우회 중.)

---

## 4. 📐 Handlebars / item 규칙
- ✅ `{{변수}}` 치환은 **HTML 탭**에서 됨(요소 내용·**비-style 속성값** OK — 예: `href="{{link}}"` `src="{{img}}"` `data-x="{{x}}"`).
- 🚫🚨🔴 **아임웹 편집기는 패널값을 바꿔도 JS를 재실행하지 않는다(2026-07-07 실측 확정)**. 그래서 **패널로 조절하는 "스타일 값"을 JS로 `el.style.xxx`에 심으면**: ①편집기에서 값 바꿔도 안 바뀜(JS 안 돎) ②처음 로드 때 박힌 **인라인 style이 잔류해 이후 CSS를 덮어씀** → 사장님이 값 바꿔도 그대로. (n09 가로폭 1200 안 커진 사건.) **✅ 정답 = 패널 스타일값은 `data-*` 속성 + CSS 속성 선택자(paint-time)로 처리**(discrete 프리셋): 예) `data-maxw="{{maxWidth}}"` + `.mm[data-maxw="1200"]{max-width:1200px}`, `data-square="{{square}}"` + `[data-square="true"]{border-radius:0}`, `data-theme` + `[data-theme="라이트"]`. **편집기에서 즉시 반영·깜빡임 0·잔류 인라인 없음.** JS로 스타일을 심는 건 **정말 연속값이라 프리셋이 불가능할 때만** 최후수단(그땐 편집기 미리보기엔 반영 안 됨을 감수).
- 🚫🚨 **CSS 값이 되는 자리의 `{{변수}}`는 전부 저장거부** — ①CSS 탭 `border-radius:{{radius}}` ②**인라인 `style="...{{}}..."`**(예: `style="border-radius:{{radius}}"` `style="aspect-ratio:{{ratio}}"`)까지 **모두 "유효하지 않은 코드"로 거부**(2026-07-06~07 실측). 아임웹 CSS 검증기가 `{{ }}`를 잘못된 CSS 토큰으로 판단. → 위 규칙대로 **`data-*` + CSS 속성 선택자**로. CSS 탭엔 정적 기본값만. JS는 값 바꿔도 재실행 안 되므로 스타일 적용 용도로 의존 금지.
- 🚫🚨 **삼중괄호 `{{{변수}}}` 금지 = "유효하지 않은 코드"로 저장거부(2026-07-05 실측)**. 표준 Handlebars의 raw 출력 문법이지만 **아임웹은 미지원**. → 항상 `{{변수}}` 이중괄호만.
- ✅ **`text-editor` 필드는 `{{변수}}` 이중괄호로도 서식 HTML(문단·줄바꿈 `<p>`)이 그대로 렌더됨**(아임웹이 리치텍스트를 escape 안 함). 즉 raw 출력하려고 삼중괄호 쓸 필요 **없음**. 문단 줄간격은 CSS `.클래스 p{margin:0 0 .15em}`로 통일.
- ✅ `{{#if}}`/`{{#unless}}` = HTML class 토글 용도 OK. (CSS 탭 분기는 1회 컴파일이라 ❌)
- 🚫 **금지(저장거부/렌더실패)**: `{{@index}}` `{{@first/last/key/root}}` `{{../부모변수}}` `{{this.field}}`. → **쓰지 말 것**(우리 회피 완료: 인덱스는 CSS counter, 부모값은 data-* 등).
- **item 한도**: 부모 **1개** · 자식 **20개** · 자식당 속성 **20개**. **중첩 item 불가.**
- 패널 1depth 항목 **최대 50개**.

---

## 5. 📍 배치 / 렌더 / 한도
- **삽입 가능**: 상단 메뉴 본문 · 글로벌 메뉴 본문 · 일반 페이지 본문. **불가**: 상단 편집영역 · **모바일 전용 섹션** · 하단 메뉴 · 모달.
- **상품 상세 개별 적용 ❌**(내부 상품데이터 접근 불가). → 상품별 위젯은 불가, 공통 또는 별도 랜딩+링크.
- 기본 **12컬럼 100% width**. 풀블리드는 위젯 CSS 트릭(가이드 미보장).
- 높이: CSS `min-height`/`height` 적용됨(일반 CSS) → **높이 옵션은 data-height+JS로**(우리 d01/d02/14/t07 적용).
- **PC/모바일 DOM 복제 안 함**(코드위젯과 다름) → **단일 렌더 + 미디어쿼리** 공식 권장. (offsetParent 트릭 불필요)
- 같은 위젯 여러 개 OK(상태공유는 안 됨).
- 한도: 계정당 위젯 **100개** · 사이트당 **100개** · 위젯당 수정 **100회**(되돌리기 없음). 탭당 **1만 자**(가이드 기준, 에디터 15,000 표시 — 보수적으로 1만 이하 유지).

- 이미지: **100MB·8000px 이하, mp4 제외**. 동영상은 `<video>` 태그 명시금지 아님(미검증)·유튜브는 썸네일+링크아웃(우리 n11).

---

## 6. 🔧 기존 43개 위젯 반영 액션 (이 가이드 적용)
1. **color 기본값 8-hex화**: d01 `rgba(26,26,26,0.5)` → `#1A1A1A80`. (color-picker 쓰는 위젯 전부)
2. **옵션버튼 `segment` 테스트**: d01의 한 필드로 `@type segment` 시도 → 진짜 버튼으로 뜨면 18개 옵션필드 일괄 전환(텍스트입력 → 버튼). 안 되면 현행 텍스트+라벨 유지.
3. `document.addEventListener('click'/'keydown')`(t03·t06)·`DOMContentLoaded`(전체) — 실측상 저장되므로 유지. 추후 저장오류 보고 시 window/element로 교체.
4. 신규 제작 시 본 가이드 1~5 항목 **무조건 준수**.

> 출처: 아임웹 FAQ q=72611/q=72610 + "프롬프트 코어" 답변(2026-06-30) + MAMORU 실측. 상충 시 **실측 > FAQ > 프롬프트코어** 우선.
