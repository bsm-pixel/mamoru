# MAMORU Brand Guide (Full Reference)

> ✅ 이 문서는 **참고용(Reference)** 입니다.
>
> - AI 작업 규칙(정본): `MAMORU/CLAUDE.md` + `.claude/ADDENDUM_*`
> - 이 문서는 브랜드/기획/컨텍스트를 길게 보관하기 위한 자료이며,
>   실제 구현 규칙/워크플로/출력 포맷은 정본을 따릅니다.

---

## Project Identity
- High-End Scissor Consulting Platform
- Stack Summary(현행 기준): Imweb(Frontend) + GAS doPost(e)(수신/기록) + Google Sheets(DB) + AppSheet(후속 관리/자동화)

---

## 1) 브랜드 페르소나 & 핵심 가치 (The MAMORU Spirit)
- **Identity:** 마모루는 단순한 미용가위 '판매자'가 아닙니다. 디자이너의 도구를 책임지는 **'기술 파트너(Tech Partner)'**입니다.
- **Slogan:** "CUT THE FAKE, KEEP THE REAL"
- **Emotional Journey:** `탐색` → `정렬` → `안심` → `행동 준비` → `망설임 제거`
- **Goal:** 고객에게 구매를 강요하지 않습니다. 대신 '여기서 결정해도 괜찮다'는 **신뢰의 상태**를 구축하는 것이 최종 목표입니다.

---

## 2) 디자인 원칙: Modern · Crisp
- **Concept:** 단순히 블랙/화이트 미니멀리즘이 아닙니다. **"클릭 가능한 것 vs 배경"이 즉시 구분되는 역할의 명확성**이 핵심입니다.
- **UI/UX Rules:**
  - **Hierarchy:** 한 화면에 두 개 이상의 강한 메시지를 두지 않아 시선 분산을 막습니다.
  - **Spacing:** 여백은 미적 요소가 아니라 **'질서의 신호'**로 사용합니다.
  - **Visual Tokens:**
    - **Base:** 다크 프리미엄 `#1C1C1E`
    - **Background:** `#FFFFFF`, `#F5F5F7`
    - **Point:** 골드 `#C9A962`는 오직 '신뢰/확정' 포인트에만 제한적으로 사용
  - **Consistency:** 위젯 간 CSS 충돌 방지를 위해 **`.mm-` 접두사**를 사용합니다.

---

## 3) 시스템 아키텍처 & 로직 (Technical Flow)
- **Environment:** Imweb (Iframe/Code Widget) + GAS doPost(e)(수신/기록) + Google Sheets(DB) + AppSheet(후속 관리/자동화)
- **진단-상담 통합 로직:**
  - **Data Flow:** 진단 페이지에서 생성된 데이터는 상담 접수 시 원활하게 이관되어야 합니다.
  - **Hybrid Case:** '진단 없는 상담'도 가능해야 하며, '진단 후 상담' 시에는 시트에 진단 결과가 포함되어 저장되어야 합니다.
- **상담/예약 유형별 처리:**
  - **매장 방문:** 1시간 단위 슬롯 예약 시스템.
  - **출장 요청:** 관리자 승인 기반 조율. 확정 시 `이동 시간 + 상담 시간`을 계산하여 전후 일정 차단(후속 관리는 AppSheet 우선).
  - **카톡 상담:** 진입 장벽을 최소화한 가장 가볍고 단순한 구조 유지.

---

## 4) 구현 가이드라인 (Implementation)
- **Performance:** 로고는 `SVG`, 애니메이션은 `Lottie` 지향.
  - 무거운 이미지 배너보다 **CSS/Code 위젯 기반 구현**을 우선하여 속도를 최적화합니다.
- **Data-Driven:** 콘텐츠(Text/Image)와 디스플레이 로직(Code)을 분리합니다.
  - 코드 수정 없이 **구글 시트 데이터 변경만으로 운영**이 가능하도록 설계합니다.
- **Responsive & Balanced Design:**
  - **Mobile-First:** 모바일 사용자를 최우선으로 고려합니다.
  - **PC Layout:** 1200px+ 환경에서도 심미성이 무너지지 않도록 반응형을 유지합니다.
  - **Container:** PC 화면에서 과도한 확장을 막기 위해 `max-width`와 중앙 정렬(`margin: 0 auto`)을 사용합니다.
- **Interactive:**
  - 기능어(상담 접수)보다 **동행/책임**이 느껴지는 마이크로카피 사용.
  - CTA는 과하지 않되 클릭 가능성은 즉시 인지되게 설계.

---

## 5) (Reference) AI 파트너 방향성
> 실제 AI 작업 규칙/역할/출력 포맷은 `CLAUDE.md` 및 `.claude/ADDENDUM_*`를 정본으로 따릅니다.

- **Mindset:** 코드를 작성할 때 항상 자문하십시오. **"이 변경이 고객의 불안을 제거하고 확신을 주는가?"**
- **Communication:** 기능 나열/가격 강조를 피하고, **구조(Structure)와 흐름(Flow)**으로 고급스러운 태도를 전달합니다.

---

## Appendix: English Context Notes (Full)

### Purpose & context
The CEO of MAMORU leads a premium professional hair scissors brand targeting hairstylists in their 20s–40s. The philosophy centers on “CUT THE FAKE, KEEP THE REAL” and emphasizes being technical experts who provide ongoing support rather than just selling products. The positioning is “Warm Premium + Approachable Expert,” combining luxury aesthetics with approachable professionalism.

The primary objective is building a comprehensive website experience focused on confidence-building through the emotional flow: “exploration → organization → reassurance → action readiness → hesitation removal.” Success is measured by customers feeling “I want to buy from this brand” rather than “they’re trying to sell to me.”

Key website components include product categorization systems, diagnostic tools for selection, consultation booking systems, and after-sales support integration. The technical setup uses Imweb with custom code widgets/iframes and backend integration via Apps Script/Sheets, with operations and automations handled primarily through AppSheet.

### Current state
Development involves multiple interconnected pages with a focus on responsive design for PC and mobile. Recent work includes refining a diagnostic tool that guides customers through career stage and cutting preferences for tailored recommendations, and improving consultation booking pages to match backend structure.

The consultation system supports three types: store visits (calendar-based scheduling), field service requests (address input + admin approval), and chat consultations via KakaoTalk. A prior misalignment between assumed features and actual implementation highlighted the importance of validating existing behavior before making changes.

### Key learnings & principles
“Modern · Crisp” has become a core design philosophy: making it immediately clear “what can be touched and what is background.” This guides hierarchy decisions and interaction design.

Avoiding banner fatigue and sales pressure is critical. Korean beauty professionals respond better to trust-building through demonstrated expertise rather than aggressive conversion tactics.

A data-driven architecture separating content from display logic enables iterative improvements without restructuring code. Seamless integration and conditional branching are essential for continuity.

### Approach & patterns
Development follows a widget-based approach with strict namespace rules to prevent conflicts. Each widget supports the customer journey while maintaining a consistent identity.

Performance optimization prioritizes SVG over PNG, inline code over image widgets, and mobile-first responsive layouts. Visual consistency relies on spacing, typography (Pretendard), and the color system (dark base #1C1C1E with limited gold accents #C9A962).

### Tools & resources
Primary platform: Imweb with HTML/CSS code widgets. Backend: Apps Script + Google Sheets (data intake and storage), with AppSheet for downstream operations.

Design uses Figma for icons and planning. Animations use Jitter for Lottie, hosted via LottieFiles.

Brand asset specs (reference): category icons at 112px canvas (2x), product cards at 900×480px, trust banners at 400×280px.