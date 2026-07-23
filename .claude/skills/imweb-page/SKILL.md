---
name: imweb-page
description: >-
  아임웹/GitHub Pages 고객 대면 페이지(히어로·상세·상담·복원수리·이벤트·재고판매 등)를 만들거나 고칠 때의
  표준(Page Kit, iframe 리사이저, 모바일 여백, 상품상세 inline 규칙)을 로드. Use when creating or
  editing any customer-facing HTML page, code widget, iframe embed, or product detail on imweb / page.mamoru.kr.
when_to_use: >-
  페이지, 아임웹, 코드위젯, 커스텀위젯, iframe, 임베드, 히어로, 상세페이지, 상품상세, 링크페이지,
  모바일 여백, page kit, 반응형, page_, 시안
---

# 아임웹·고객 페이지 제작 표준

> 고객 대면 = 브랜드 100% 준수. 디자인 기준은 `mamoru-brand` 스킬 / Brand Guide 병행.

## 1. Page Kit = 전 페이지 통합 표준
- 복사 원본 `projects/_design_lab/_page_kit.html`.
- 구조: 풀블리드 배경 `.mm-band` > 중앙 `.mm-inner`(**760** / 와이드 **1100**). 음수마진·`100vw` 해킹 폐기.
- 변수 통일 `--void/--cream`. 히어로 라이트 기본, 다크(Void)는 몰입 띠만.

## 2. iframe 임베드 (page.mamoru.kr → 아임웹)
- 아임웹은 **코드위젯 안 `<script>`를 제거/미실행**한다. 높이조절을 자식 스크립트에 의존하면 800px에서 잘림.
- 해결: 자식은 `postMessage({type:'MAMORU_IFRAME_SIZE', id, height})` 전송(monotonic·REQUEST_HEIGHT 재요청). 실제 리사이저는 **전역 헤더 `banner-widget.js`**(source 매칭). **공통코드 무수정.**
- 임베드 판별: `window.parent!==window` → `html.mm-embed`. 고정 CTA/전체화면 레이어는 임베드일 때 absolute 앵커로.
- 표준 패턴 참조: memory `reference_iframe_codewidget_pattern`, `reference_imweb_strips_codewidget_script`.

## 3. 커스텀 위젯 (아임웹 디자인모드)
- **fetch/XHR·외부 라이브러리·iframe·localStorage·내부데이터 전부 금지** (fetch=저장거부). → 라이브 데이터 위젯 ❌, 자기완결 인터랙션만 ✅.
- 🔴 **CSS값 자리 `{{ }}` 저장거부** → 동적 스타일은 `data-*` + CSS 속성선택자. 편집기는 패널값 바꿔도 JS 재실행 안 함 → JS로 `el.style` 심지 말 것.

## 4. 모바일 (제일 자주 깨지는 곳)
- **좌우 여백 20px 표준**: 히어로 아닌 콘텐츠·배너는 모바일 좌우 20px 필수. 가로영역 100% 확장이 기본.
- 가로스크롤 여백은 **트랙**에(컨테이너 padding-right는 스크롤 끝에 안 먹음). scroll-snap이면 `scroll-padding-left:20px`.
- 🚨 **헤드리스 `--window-size`만으로 모바일 열수 판정 금지** — vw·auto-fit 실제와 다름. puppeteer `setViewport({isMobile:true})` + `getComputedStyle` 실측. inline은 `auto-fit minmax`+clamp(고정 repeat 금지).

## 5. 아임웹 상품 상세 본문
- **inline style 전용**: `<style>/<script>/<iframe>` 차단. inline + `clamp()`만.
- 서체: 영문숫자=Paperlogy/한글=Pretendard (스택 아닌 **span 분리**). 폰트는 아임웹 전역헤더에서만 로드.

## 6. 검증·배포
- 실제 렌더 확인(Chrome 헤드리스 캡처, PC+모바일 폭). 자산 `src/href` 실제 파일 매칭 확인(깨진 이미지=고객 노출).
- 배포 후 `curl`로 라이브 콘텐츠·해시 대조(이미지 교체는 Fastly 엣지캐시 함정 — `git hash-object` 3해시 대조).
- 🚫 `common_code/header_code.txt` 등 공유 공통코드 절대 손대지 않기.
- 리뷰 표시 수정은 `page_reviews.html` + `page_main_btm.html` **2곳 동시**.

## 폴더
consulting/as/main/event/products/reviews/stock_sale = 고객 HTML(page.mamoru.kr). 🚫 폴더명 변경 금지(알림톡·위젯 박제).
허브: `projects/_design_lab/index.html` (전체 페이지 인덱스).
