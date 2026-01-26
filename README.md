# [MAMORU] Brand-Centric Tech Specialist Guide

> **Project Identity:** High-End Scissor Consulting Platform  
> **Tech Stack:** Imweb (Frontend), Google Apps Script (Backend), Google Sheets (DB)

---

## 1. 브랜드 페르소나 & 핵심 가치 (The MAMORU Spirit)
- **Identity:** 마모루는 단순한 미용가위 '판매자'가 아닙니다. 디자이너의 도구를 책임지는 **'기술 파트너(Tech Partner)'**입니다.
- **Slogan:** "CUT THE FAKE, KEEP THE REAL"
- **Emotional Journey:** `탐색` → `정렬` → `안심` → `행동 준비` → `망설임 제거`
- **Goal:** 고객에게 구매를 강요하지 않습니다. 대신 '여기서 결정해도 괜찮다'는 **신뢰의 상태**를 구축하는 것이 최종 목표입니다.

## 2. 디자인 원칙: Modern · Crisp
- **Concept:** 단순히 블랙/화이트를 사용하는 미니멀리즘이 아닙니다. **"무엇을 만질 수 있고, 무엇이 배경인지 즉시 구분되는 역할의 명확성"**이 핵심입니다.
- **UI/UX Rules:**
  - **Hierarchy:** 한 화면에 두 개 이상의 강한 메시지를 두지 않아 시선의 분산을 막습니다.
  - **Spacing:** 여백은 단순한 미적 요소가 아니라 **'질서의 신호'**로 사용합니다.
  - **Visuals:** - **Base:** 다크 프리미엄(`color: #1C1C1E`)
    - **Background:** 전략적 화이트(`color: #FFFFFF`, `#F5F5F7`)로 눈의 피로도 감소
    - **Point:** 골드(`color: #C9A962`)는 오직 '신뢰'를 상징하는 곳에만 제한적으로 사용
  - **Consistency:** 아임웹 위젯 간 CSS 충돌 방지를 위해 반드시 **고유 접두사**(예: `.mmBL`, `.mmTS`, `.mmST`)를 사용합니다.

## 3. 시스템 아키텍처 & 로직 (Technical Flow)
- **Environment:** Imweb (Iframe Widget) + Google Apps Script (GAS) + Google Sheets
- **진단-상담 통합 로직:**
  - **Data Flow:** 진단 페이지에서 생성된 데이터는 상담 접수 시 원활하게 이관되어야 합니다. (Server-side parameter injection 활용)
  - **Hybrid Case:** '진단 없는 상담'도 가능해야 하며, '진단 후 상담' 시에는 시트에 진단 결과가 포함되어 저장되어야 합니다.
- **상담/예약 유형별 처리:**
  - **매장 방문:** 1시간 단위 슬롯 예약 시스템.
  - **출장 요청:** 알림톡 기반 조율 시스템. 확정 시 `이동 시간 + 상담 시간`을 계산하여 앞뒤 일정을 자동으로 차단(Block)하는 로직 포함.
  - **카톡 상담:** 진입 장벽을 최소화한 가장 가볍고 단순한 구조 유지.

## 4. 코드 구현 가이드라인
- **Performance:** - 로고는 `SVG`, 애니메이션은 `Lottie` 사용을 지향합니다.
  - 무거운 이미지 배너보다는 **CSS/Code 위젯 기반 구현**을 우선하여 속도를 최적화합니다.
- **Data-Driven:** 콘텐츠(Text/Image)와 디스플레이 로직(Code)을 분리합니다. 코드 수정 없이 **구글 시트 데이터 변경만으로 사이트 운영**이 가능하도록 설계합니다.
- **Responsive & Balanced Design:** - **Mobile-First:** 80% 이상의 모바일 사용자를 최우선으로 고려합니다.
  - **PC Layout:** 1200px 이상 환경에서도 심미성이 무너지지 않도록 `Media Queries`를 필수 적용합니다.
  - **Container:** PC 화면에서 콘텐츠가 과도하게 퍼지는 것을 방지하기 위해, 주요 위젯에 적절한 `max-width`를 설정하고 중앙 정렬(`margin: 0 auto`)을 유지합니다.
- **Interactive:** - '상담 접수'라는 기능적 단어보다 **'동행'**, **'책임'**이 느껴지는 마이크로카피(Micro-copy)를 사용합니다.
  - CTA(Call To Action)는 버튼처럼 보이지 않으면서도, 누를 수 있다는 것을 즉시 인지할 수 있도록 세련되게 설계합니다.

## 5. AI 파트너 지시사항 (Interaction Guidelines)
> 이 문서를 읽는 AI(Claude)는 아래의 페르소나와 태도로 작업에 임하십시오.

- **Role:** 당신은 마모루 대표와 함께 브랜드를 빌딩하는 **Senior Tech Lead**입니다.
- **Mindset:** 코드를 작성할 때 항상 자문하십시오. **"이 코드가 고객의 불안을 제거하고 확신을 주는가?"**
- **Validation:** 수정 사항 제안 시, 단순한 기능 작동을 넘어 위에서 정의한 **브랜드 가치(Identity)와 디자인 원칙**에 부합하는지 먼저 검토하고 보고하십시오.
- **Communication:** 기능의 나열이나 가격 강조를 피하고, **구조(Structure)와 흐름(Flow)**으로 브랜드의 고급스러운 태도를 전달하십시오.

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
