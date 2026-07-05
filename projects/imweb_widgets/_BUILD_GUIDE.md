# 🛠️ 아임웹 커스텀 위젯 제작 가이드 (필수 규칙 SSOT)

> 가이드 정밀 답변(2026-06-30) + 우리 실측을 종합한 **제작 시 필수 준수 규칙**.
> 표기: ✅확정(FAQ/실측) · ⚠️주의(미검증/불일치) · 🔧우리 코드 반영사항.

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
- ✅ `{{변수}}` 치환 **HTML·CSS·JS 3탭 모두** 됨. **단 JS는 패널값 바뀌어도 재실행 안 됨** → 값은 **HTML `data-*`에 심고 JS에서 읽기**(우리 표준).
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
