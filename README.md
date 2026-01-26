# [MAMORU] Brand-Centric Tech Specialist Guide

## 1. 브랜드 페르소나 & 핵심 가치 (The MAMORU Spirit)
- **Identity:** 마모루는 미용가위를 파는 '판매자'가 아니라, 디자이너의 도구를 책임지는 '기술 파트너'다.
- **Slogan:** "CUT THE FAKE, KEEP THE REAL"
- **Emotional Journey:** 탐색 → 정렬 → 안심 → 행동 준비 → 망설임 제거.
- **Goal:** 고객에게 '사라'고 강요하지 않는다. 대신 '여기서 결정해도 괜찮다'는 신뢰의 상태를 만든다.

## 2. 디자인 원칙: Modern · Crisp
- **해석:** 단순히 검정/흰색을 쓰는 미니멀리즘이 아니다. "무엇을 만질 수 있고, 무엇이 배경인지 즉시 구분되는 역할의 명확성"이 핵심이다.
- **UI/UX 규칙:**
  - **Hierarchy:** 한 화면에 두 개 이상의 강한 메시지를 두지 않는다.
  - **Spacing:** 간격은 미적 요소가 아니라 '질서의 신호'다.
  - **Visuals:** 다크 프리미엄(#1C1C1E)과 전략적 화이트(#FFFFFF, #F5F5F7)를 조합하여 피로도를 낮춘다. 골드 포인트(#C9A962)는 신뢰의 상징으로만 사용한다.
  - **Consistency:** 아임웹 위젯 간 충돌 방지를 위해 반드시 고유 접두사(예: .mmBL, .mmTS, .mmST 등)를 사용한다.

## 3. 시스템 아키텍처 & 로직 (Technical Flow)
- **환경:** Imweb(Iframe) + Google Apps Script(GAS) + Google Sheets.
- **진단-상담 통합 로직:**
  - 진단 페이지에서 생성된 데이터는 상담 접수 시 원활하게 전송되어야 한다. (Server-side parameter injection 활용)
  - 진단 없이도 상담이 가능해야 하며, 진단 후 상담 시에는 시트에 진단 데이터가 포함되어 저장되어야 한다.
- **상담/예약 유형별 처리:**
  - **매장 방문:** 1시간 단위 예약 시스템.
  - **출장 요청:** 알림톡 기반 조율 시스템. 확정 시 '이동 시간 + 상담 시간'을 고려하여 앞뒤 일정을 자동으로 차단(Block)하는 로직을 포함한다.
  - **카톡 상담:** 가장 가볍고 단순한 진입 구조 유지.

## 4. 코드 구현 가이드라인
- **Performance:** 로고는 SVG, 애니메이션은 Lottie(LottieFiles), 이미지 보다는 코드 위젯 기반 구현을 우선한다.
- **Data-Driven:** 콘텐츠와 디스플레이 로직을 분리하여 코드 수정 없이 데이터(시트) 변경만으로 운영이 가능하게 짠다.
- **Responsive & Balanced Design:** 80% 이상의 모바일 사용자를 고려한 Mobile-First 기법을 기본으로 하되, PC 환경(1200px 이상)에서도 가독성과 심미성이 무너지지 않도록 **반응형 디자인(Media Queries)**을 필수 적용한다.
- **PC Layout Control:** PC 화면에서 컨텐츠가 무분별하게 양옆으로 퍼지는 것을 방지하기 위해, 주요 위젯 컨테이너에 적절한 `max-width`를 설정하고 중앙 정렬을 유지한다.
- **Font & Spacing Scalability:** 모바일의 좁은 화면과 PC의 넓은 화면 모두에서 'Modern · Crisp'한 감각이 유지되도록, 화면 크기에 따라 폰트 크기와 섹션 간격을 유연하게 조절한다.
- **Interactive:** '상담 접수'라는 단어보다 '동행'과 '책임'이 느껴지는 UI를 구성하되, 선택지(CTA)는 버튼처럼 보이지 않으면서도 즉시 인지되도록 설계한다.

## 5. 커서(AI)에게 내리는 지시 (Interaction Style)
- 당신은 마모루의 대표와 함께 브랜드를 빌딩하는 Senior Engineer다.
- 코드를 짤 때 항상 "이 코드가 고객의 불안을 제거하고 확신을 주는가?"를 자문하라.
- 수정 사항 제안 시, 단순 기능 개선을 넘어 브랜드 지침(6-1 ~ 6-6 섹션 역할)에 부합하는지 먼저 검토하고 보고하라.
- 불필요한 기능 나열이나 가격 강조를 피하고, 구조와 흐름으로 브랜드의 태도를 전달하라.

- Purpose & context
대표 is the CEO of MAMORU, a premium professional hair scissors brand targeting hairstylists in their 20s-40s. The core business philosophy centers on "CUT THE FAKE, KEEP THE REAL" and emphasizes being technical experts who provide ongoing support rather than just selling products. The brand positioning is "Warm Premium + Approachable Expert" - combining luxury aesthetics with approachable professionalism to build trust through expertise rather than aggressive sales tactics.
The primary objective is developing a comprehensive brand website that creates a customer experience focused on confidence-building through the emotional flow: "exploration → organization → reassurance → action readiness → hesitation removal." Success is measured by customers feeling "I want to buy from this brand" rather than "they're trying to sell to me."
Key website components include product categorization systems, diagnostic tools for scissors selection, consultation booking systems, and after-sales support integration. The technical setup uses Imweb platform with custom code widgets, iframe structures, and backend integration through Apps Script and Google Sheets.
Current state
대표 is actively developing multiple interconnected website pages with a focus on responsive design for both PC and mobile. Recent work includes refining a diagnostic tool that guides customers through questions about career stage and cutting preferences to provide customized scissors recommendations, and improving the consultation booking page to properly match the backend system structure.
The consultation system currently supports three types: store visits (with calendar-based scheduling), field service requests (requiring address input and admin approval), and chat consultations via KakaoTalk. Recent discovery revealed misalignment between assumed system features and actual implementation, requiring design corrections.
Key learnings & principles
The "Modern · Crisp" design philosophy has emerged as core to the brand identity, defined as making it immediately clear "what can be touched and what is background." This principle guides all visual hierarchy decisions and user interface design.
A critical insight is avoiding "banner fatigue" and sales pressure through varied design elements and natural customer flow patterns. The brand has learned that Korean beauty industry professionals respond better to trust-building through expertise demonstration rather than aggressive conversion tactics.
Technical architecture benefits from data-driven design that separates content from display logic, making it easy to modify questions and flows without restructuring code. Conditional branching and seamless system integration have proven essential for user experience continuity.
Approach & patterns
Development follows a systematic widget-based approach using unique CSS class prefixes to prevent conflicts when multiple components are used together. Each widget serves specific functions in the customer journey while maintaining consistent visual identity.
Design decisions prioritize performance optimization, favoring SVG over PNG for logos, inline code over image widgets for speed, and mobile-first responsive design. Visual consistency is maintained through careful attention to spacing, typography (Pretendard font family), and color systems featuring dark backgrounds (#1C1C1E) with gold accents (#C9A962).
The workflow involves iterative refinement based on user experience testing, with frequent adjustments to visual hierarchy, spacing, and interaction patterns. Strategic use of "breathing space" between product sections prevents overwhelming customers with repetitive displays.
Tools & resources
Primary development platform is Imweb with custom HTML/CSS code widgets. Backend integration uses Apps Script and Google Sheets for consultation management and data processing. Design work utilizes Figma for icon creation and visual planning.
Animation implementation leverages Jitter for Lottie animations with LottieFiles hosting for web deployment. The technical stack includes iframe structures for diagnostic tools and URL parameter passing for seamless system integration between different page components.
Brand assets follow specific technical specifications: category icons at 112px canvas size for 2x resolution display, product card images at 900x480px, and trust banner images at 400x280px for optimal mobile performance.
