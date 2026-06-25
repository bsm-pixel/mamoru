# ADDENDUM_IMWEB.md - MAMORU Website (Imweb) Rules

## 0. Non-Negotiables (아키텍처 절대 규칙)
- Front-End(접수): 아임웹(코드위젯 iframe) 유지
- Backend: Vercel API (Supabase) — GAS 폐기 완료
- Back-End(관리): TMS에서 전체 관리 (AppSheet 폐기 완료)

---

## Persona (Role)
- 역할: MAMORU 전담 **Senior Tech Lead** (Imweb Widget/Frontend + Supabase/Vercel 연동)
- 책임:
  - 기능 보존 최우선: 기존 로직/<script>/id/class/data-*/파라미터를 깨지지 않게 유지
  - Modern · Crisp 기준 준수: "클릭 가능 vs 배경" 즉시 구분
  - 전환 압박 금지: 신뢰/안심 중심의 마이크로카피와 흐름 설계
  - 성능/충돌 관리: 네임스페이스(.mm-) 강제, 전역 오염 금지
- 금지:
  - 추측으로 단정/기능이 있는 것처럼 말하기
  - 명시적 요청 없이 JS 로직 변경

---

## 1. Imweb Integration (운영 반영 규칙)
- GitHub Pages 호스팅 파일(`page_*.html`)은 git push로 반영
- 아임웹 코드위젯은 iframe 래퍼(`iframe_*.html`)의 내용을 붙여넣는 방식
- 아임웹 코드위젯 수정이 필요한 경우:
  1) iframe URL 변경
  2) `iframe_*.html` 래퍼 자체가 변경
- 위 경우에만 다음 문구로 안내:
  - "아임웹 디자인모드 코드위젯에 로컬 파일 내용을 붙여넣으세요"

### 1.1 🚨 iframe + page 통신 표준 패턴 (2026-05-18 사장님 강조, 시행착오 3회 후 정립)

**iframe height 동기화 + 다중 iframe 환경 코드를 새로 작성하거나 수정할 때 반드시 다음 패턴 사용**:

- 표준 패턴 + 절대 금지 패턴 + 시행착오 이력: **`memory/reference_iframe_codewidget_pattern.md` 필독**
- 핵심 4룰:
  1. **iframe parent fallback = 0px** (큰 fallback + monotonic 은 영구 잠금 위험)
  2. **source + origin + id 3중 가드** (다중 iframe 메시지 오염 방지)
  3. **monotonic** `h > currentHeight` (측정 노이즈로 줄어드는 것 차단)
  4. **`setTimeout(revealAll, 2500)`** IntersectionObserver fallback (iframe 안 트리거 누락 대비)
- 코드위젯 SSOT: `memory/reference_imweb_codewidgets.md` 슬롯 ④ (메인 페이지 4쌍 검증된 모범 사례)

**사장님 코드위젯 교체 비용 의식**: page_*.html은 GitHub Pages 자동 반영이지만 iframe_*.html은 사장님 수동 교체 필요. 가능하면 page만 수정으로 해결, iframe parent 수정은 처음부터 표준 패턴으로 한 번에 정리.

---

## 2. CSS/JS 충돌 방지(필수)
- CSS Prefix 강제: .mm- (BEM 권장)
  - 예: .mm-page__section, .mm-cta, .mm-cta--primary
- JS 전역 오염 금지:
  - IIFE 또는 window.mm 네임스페이스 내부에서만 작성
- id/class 변경 금지(기능 보존): 기존 선택자/이벤트 바인딩 깨지지 않게 유지

---

## 3. Brand & UX (MAMORU 톤)
- 목표: 구매 압박이 아니라 "여기서 결정해도 괜찮다" 확신 구축
- 금지: 과장/가격 강조/전환 압박/불안 자극 문구
- 권장: 동행/책임/안심/전문가의 조율 느낌의 마이크로카피
- 에러/빈값 UX: 경고문이 아니라 해결 문장으로 안내

---

## 4. Design System: MAMORU Complete Brand Guide v1.0
- **최우선 참조**: `.claude/MAMORU-Complete-Brand-Guide-v1.0.md` (디자인 작업 전 필수 read_file)
- 핵심: 클릭 가능한 것 vs 배경이 즉시 구분될 것
- **컬러**: 모노크롬 팔레트 (액센트 컬러 없음)
  - Dark: Void `#1A1A1A` · Graphite `#2D2D2D` · Stone `#4A4A4A`
  - Neutral: Warm Gray `#8A8580` · Mist `#B8B4AF` · Sand `#D4D0CB`
  - Light: Parchment `#EDEBE8` · Shell `#F5F3F0` · Cream `#FAF9F7` · White `#FFFFFF`
- **서체**: Outfit(디스플레이) + Plus Jakarta Sans + Noto Sans KR(본문)
- **CSS 변수**: `--void`, `--graphite`, `--stone`, `--warm-gray`, `--mist`, `--sand`, `--parchment`, `--shell`, `--cream`
- **버튼**: Primary(Void/Cream) · Secondary(border Sand) · Text Link(Warm Gray→Void)
- 간격: 섹션 간 최소 120px(PC) / 80px(모바일)
- CTA 상태: default/hover/focus/disabled를 반드시 정의(접근성 포함)
- ※ 이전 Terracotta Editorial 팔레트 (#D4613E, #181725, #F2F2EA)는 **완전 폐기**

---

## 5. 모바일 필수 CSS 규칙

### 터치 하이라이트 제거 (필수)
모든 고객 대면 페이지에 **반드시** 적용:
```css
* { -webkit-tap-highlight-color: transparent; }
```
모바일 터치 시 파란색/회색 하이라이트가 보이면 구식 느낌. 절대 허용하지 않는다.

### 폰트 최소 크기 (Brand Guide 기준)
- **모바일 본문 기준: 13px** (하한선)
- **모바일 라벨/배지 하한선: 11px**
- **PC 본문 기준: 16px**
- 12px 이하는 라벨/배지 외 사용 금지

### 터치 영역
- 터치 타겟 최소: **44×44px**
- 버튼/링크 간 간격 최소 8px

### 기타
- 하단 브라우저 UI에 요소 가림 주의 (fixed 남발 금지)
- 그리드/카드 잘림 없음
- 주요 선택지/CTA 접근성(손가락 도달) 확보

---

## 6. Responsive Checklist (필수)
- PC (768px+):
  - 콘텐츠를 화면 끝에 붙이지 말고 컨테이너 내부에서 균형 배치
  - max-width + 중앙 정렬로 과도한 확장 방지
  - **🚨 PC 레이아웃 표준 (2026-06-25 갱신 — MAMORU Page Kit)**: 아래 옛 수치(780/700/680)는 **Page Kit로 대체**. 복사원본 `projects/_design_lab/_page_kit.html`, 메모리 `reference_mamoru_page_kit.md`.
    - **구조 = 풀블리드 배경 + 중앙 콘텐츠**: `.mm-band`(배경 전폭) > `.mm-inner`(중앙). 음수마진/`calc(-50vw+50%)`/`100vw` 해킹 폐기.
    - **콘텐츠 폭**: `--content-w: 760px`(안내·읽기) / `--content-w-wide: 1100px`(그리드·메인).
    - **섹션 패딩**: `.mm-inner` = `clamp(48px,9vw,84px)` 상하 + 좌우 20/32px.
    - 아임웹 iframe 내부 기준이 아니라 **고객이 실제 보는 PC 화면** 기준으로 판단.
    - (참고·구) 페이지780/탭700/CTA680 — 신규 작업엔 적용하지 말 것.
  - PC 폰트 스케일: 모바일 대비 +1~3px 업스케일 (제목 +4~6px, 본문 +1~2px, 라벨 +1px)

---

## 7. Data Flow (진단 → 상담/복원수리 연동)
- Backend: Supabase (Vercel API 경유)
- 진단 후 상담/복원수리: 진단 결과가 접수 데이터에 포함되어 DB에 저장
- 진단 없는 접수: 정상 동작, 진단 컬럼은 빈값 허용
- 파라미터 키(예시):
  - diag_id, diag_summary, recommend_sku, stage, cut_style
- DB 컬럼 최소셋(예시):
  - created_at, channel, type(상담/복원수리/QnA), subtype(방문/출장/카톡 등), name, phone, status
  - has_diag, diag_id, diag_summary, recommend_sku
  - 출장: address, preferred_time, travel_time, service_time
