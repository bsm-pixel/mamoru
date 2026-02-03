# [MAMORU] Brand-Centric Tech Specialist Guide

> **Project Identity:** High-End Scissor Consulting Platform  
> **Tech Stack:** Imweb(Frontend) + GAS doPost(e)(수신/기록) + Google Sheets(DB) + AppSheet(후속 관리/자동화)

---

## 작업 규칙(정본)
- AI 작업 규칙의 정본은 `CLAUDE.md` 및 `.claude/ADDENDUM_*` 입니다.
- 아임웹/상담/AS/QnA 작업 시: `.claude/ADDENDUM_IMWEB.md` 기준

---

## 1. 브랜드 페르소나 & 핵심 가치 (The MAMORU Spirit)
- **Identity:** 마모루는 단순한 미용가위 '판매자'가 아닙니다. 디자이너의 도구를 책임지는 **'기술 파트너(Tech Partner)'**입니다.
- **Slogan:** "CUT THE FAKE, KEEP THE REAL"
- **Emotional Journey:** `탐색` → `정렬` → `안심` → `행동 준비` → `망설임 제거`
- **Goal:** 고객에게 구매를 강요하지 않습니다. 대신 '여기서 결정해도 괜찮다'는 **신뢰의 상태**를 구축하는 것이 최종 목표입니다.

## 2. 디자인 원칙: Modern · Crisp
- **Concept:** 단순한 블랙/화이트 미니멀리즘이 아니라, **“클릭 가능한 것 vs 배경”이 즉시 구분되는 역할의 명확성**이 핵심입니다.
- **UI/UX Rules:**
  - **Hierarchy:** 한 화면에 두 개 이상의 강한 메시지를 두지 않습니다.
  - **Spacing:** 여백은 미적 요소가 아니라 **‘질서의 신호’**로 사용합니다.
  - **Visual Tokens:**
    - **Base:** 다크 프리미엄 `#1C1C1E`
    - **Background:** `#FFFFFF`, `#F5F5F7`
    - **Point:** 골드 `#C9A962` (오직 ‘신뢰/확정’ 포인트에만 제한)
  - **Consistency:** 아임웹 위젯 간 CSS 충돌 방지를 위해 **`.mm-` 접두사**를 사용합니다.

## 3. 시스템 아키텍처 & 로직 (Technical Flow)
- **Environment:** Imweb (Iframe/Code Widget) + GAS doPost(e)(수신/기록) + Google Sheets(DB) + AppSheet(후속 관리/자동화)
- **진단-상담 통합 로직:**
  - **Data Flow:** 진단 결과는 상담/AS 접수 시 자연스럽게 이관되어 시트에 저장되어야 합니다.
  - **Hybrid Case:** '진단 없는 접수'도 가능해야 하며, '진단 후 접수' 시 진단 결과가 포함되어 저장되어야 합니다.
- **상담/예약 유형별 처리:**
  - **매장 방문:** 1시간 단위 슬롯 예약.
  - **출장 요청:** 주소 입력 + 관리자 승인 + 확정 시 `이동 시간 + 상담 시간` 기준으로 전후 일정 차단(후속 관리는 AppSheet 우선).
  - **카톡 상담:** 진입 장벽 최소(가장 단순한 구조 유지).

### 3.1 A/S 접수 시스템
- **위치:** `projects/as/_gas/`
- **Tech Stack:** GAS doPost + Google Sheets + MAKE(알림톡) + 롯데택배 API

#### 진행방식
| 방식 | 설명 |
|------|------|
| **방문수거** | 롯데택배 기사님이 고객 주소지로 방문하여 수거 |
| **직접발송** | 고객이 직접 택배 발송 (수거비 무료) |

#### 가격 정책
| 항목 | 단가 |
|------|------|
| 마모루 가위 A/S | 10,000원/자루 |
| 타사 가위 A/S | 20,000원/자루 (마모루 가위 접수 시에만 추가 가능) |

#### 수거비 정책 (방문수거 시)
| 총 수량 | 수거비 |
|---------|--------|
| 1자루 | 5,000원 |
| 2자루 | 3,000원 |
| 3자루 이상 | 무료 |
| 직접발송 | 무료 |

> **총 수량** = 마모루 가위 + 타사 가위 합산

#### MAKE 웹훅 Payload (AS_CREATE)
```json
{
  "as_id": "AS-20260203-001",
  "name": "고객명",
  "phone": "01012345678",
  "proceed_type": "방문수거|직접발송",
  "pickup_date": "2026-02-05",
  "qty_mamoru": "2",
  "qty_other": "1",
  "service_cost": 40000,
  "shipping_fee": 0,
  "total_amount": 40000
}
```

## 4. 파일 구조 & 배포 방식

### 4.1 프로젝트 디렉토리
```
projects/
├── common_code/                # 아임웹 공통 코드 (환경설정)
│   ├── header_code_top.txt     ← Header Code 상단 (preconnect)
│   ├── header_code.txt         ← Header Code (iframe CSS)
│   ├── header_product_style.txt ← Header Code 추가 (상품 진열 공통 스타일)
│   └── footer_code.txt         ← Footer Code (iframe JS)
├── product_page/               # 상품 진열 / 카테고리 페이지
│   └── productpage_categorybar.txt ← 카테고리 TOP 코드위젯 (배너+탭)
├── as/                         # A/S 접수
│   ├── page_form.html          ← GitHub Pages 배포 (실제 콘텐츠)
│   ├── page_guide.html         ← GitHub Pages 배포
│   ├── iframe_form.html        ← 아임웹 코드위젯에 붙여넣기 (/53)
│   ├── iframe_guide.html       ← 아임웹 코드위젯에 붙여넣기 (/60)
│   ├── icons/                  # SVG 아이콘
│   ├── holidays_2026_2027.tsv  # 공휴일 원본 데이터
│   └── _gas/                   # GAS 백엔드 (Code.gs 등)
├── consulting/                 # 상담접수
│   ├── page_form.html          ← GitHub Pages 배포
│   ├── page_diag.html          ← GitHub Pages 배포
│   ├── page_guide.html         ← GitHub Pages 배포
│   ├── page_recommend.html     ← GitHub Pages 배포
│   ├── iframe_form.html        ← 아임웹 코드위젯
│   ├── iframe_diag.html        ← 아임웹 코드위젯
│   ├── iframe_guide.html       ← 아임웹 코드위젯 (/61)
│   ├── iframe_recommend.html   ← 아임웹 코드위젯
│   ├── products.js             # 제품 데이터 (추천 로직)
│   └── Code.gs                 # GAS 백엔드
└── brand/                      # 브랜드 소개
    ├── page_intro.html         ← GitHub Pages 배포
    ├── iframe_intro.html       ← 아임웹 코드위젯
    └── Code.gs                 # GAS 백엔드
```

### 4.2 배포 구조 (3단계)

| 계층 | 파일 | 배포 방식 | 변경 시 반영 |
|------|------|-----------|-------------|
| **공통 코드** | `common_code/*.txt` | 아임웹 환경설정에 수동 붙여넣기 | **아임웹 환경설정에서 직접 교체** |
| **콘텐츠** | `page_*.html` | GitHub Pages 자동 배포 | git push → 1~2분 후 자동 |
| **코드위젯** | `iframe_*.html`, `productpage_*.txt` | 아임웹 코드위젯에 수동 붙여넣기 | **아임웹에서 직접 교체 필요** |
| **백엔드** | `Code.gs` | GAS 에디터에서 배포 | GAS 새 배포 필요 |

> **중요:** `page_*.html` 수정은 push만 하면 되지만, `iframe_*.html` 수정 시에는 **아임웹 해당 페이지의 코드위젯도 동일 내용으로 교체**해야 반영됩니다.

### 4.3 아임웹 공통 코드 매핑

| 아임웹 환경설정 영역 | 로컬 파일 | 역할 |
|-------------------|----------|------|
| Header Code 상단 | `common_code/header_code_top.txt` | preconnect/dns-prefetch |
| Header Code | `common_code/header_code.txt` + `header_product_style.txt` | iframe CSS + 상품 진열 공통 스타일 |
| Footer Code | `common_code/footer_code.txt` | iframe 높이/모달 뷰포트 JS |

> **상품 진열 스타일**은 Header Code에 한 번만 넣으면 전 카테고리 페이지에 적용됩니다. 새 상품 위젯 추가 시 `header_product_style.txt`의 `IDS` 배열에 위젯 ID만 추가하면 됩니다.

### 4.4 아임웹 페이지 매핑

| 아임웹 경로 | 코드위젯 파일 | 로드하는 콘텐츠 |
|------------|--------------|---------------|
| `/53` | `projects/as/iframe_form.html` | AS 접수 폼 |
| `/60` | `projects/as/iframe_guide.html` | AS 안내 |
| `/61` | `projects/consulting/iframe_guide.html` | 상담 안내 |
| (상담접수) | `projects/consulting/iframe_form.html` | 상담접수 폼 |
| (간편진단) | `projects/consulting/iframe_diag.html` | 간편진단 |
| (맞춤추천) | `projects/consulting/iframe_recommend.html` | 맞춤 추천 |
| (브랜드) | `projects/brand/iframe_intro.html` | 브랜드 소개 |

### 4.5 카테고리 페이지 구조

각 카테고리 페이지(블런트/틴닝/장가위/슬라이싱)는 **코드위젯 2개**로 구성:

| 순서 | 코드위젯 | 파일 | 역할 |
|------|---------|------|------|
| 1 | 카테고리 TOP | `productpage_categorybar.txt` | 배너 이미지(PC/모바일 반응형) + 텍스트 오버레이 + 블랙 칩 탭 |
| 2 | 상품 진열 | 아임웹 기본 상품 위젯 | 스타일은 공통 코드(`header_product_style.txt`)에서 일괄 적용 |

**카테고리 TOP 수정 포인트** (`productpage_categorybar.txt`):
- `<source srcset="...">` / `<img src="...">` — 배너 이미지 URL (PC/모바일)
- `.cat-banner-en` — 영문 카테고리명 (예: BLUNT)
- `.cat-banner-kr` — 한글 설명 (예: 커트 시작의 출발점)
- `.cat-mini a.active` — 현재 페이지에 해당하는 탭에 active 클래스

**카테고리 페이지 매핑:**

| 아임웹 경로 | 카테고리 | 상품 위젯 ID |
|------------|---------|-------------|
| `/29` | 블런트 | `w2025082780aa858e21dc6` |
| `/30` | 틴닝 | (위젯 ID 확인 후 IDS 배열에 추가) |
| `/39` | 장가위 | (위젯 ID 확인 후 IDS 배열에 추가) |
| `/40` | 슬라이싱 | (위젯 ID 확인 후 IDS 배열에 추가) |

### 4.6 iframe ↔ Parent 공통 메시지 프로토콜 (postMessage)

모든 `page_*.html` ↔ `iframe_*.html` 통신에 사용되는 표준 메시지 타입:

| 메시지 type | 방향 | 용도 |
|------------|------|------|
| `MAMORU_IFRAME_SIZE` | page → parent | iframe 높이 자동 조절 (`{ height }`) |
| `MAMORU_SCROLL_TOP` | page → parent | 부모 페이지 맨 위로 스크롤 |
| `MAMORU_NAVIGATE` | page → parent | 부모 페이지 이동 (`{ url }`) |
| `MAMORU_REQUEST_VIEWPORT` | page → parent | 모달 중앙 배치용 뷰포트 정보 요청 |
| `MAMORU_VIEWPORT_INFO` | parent → page | 뷰포트 정보 응답 (`{ visibleTop, visibleHeight }`) |

> `MAMORU_REQUEST_VIEWPORT` / `MAMORU_VIEWPORT_INFO`는 iframe 내부에서 모달/달력을 **현재 보이는 화면 중앙**에 배치하기 위해 사용합니다. 모달이 있는 page에서 요청하면, iframe_*.html(부모)이 응답합니다.

---

## 5. 구현 가이드라인(요약)
- **Performance:** 로고는 SVG, 애니메이션은 Lottie 지향. 이미지 배너 남발 대신 코드/스타일 위젯 우선.
- **Data-Driven:** 콘텐츠(Text/Image)와 로직(Code) 분리. 시트 변경만으로 운영 가능하게 설계.
- **Responsive:** Mobile-first. PC에서는 max-width + 중앙정렬로 과도한 확장 방지.
- **Micro-copy:** 기능 나열/가격 강조 대신, 동행/책임/안심이 느껴지는 문장 사용.

---

## 6. 변경 이력 (최근)

### 2026-02-03 (상품 진열 / 카테고리 페이지)
- **style:** 상품 진열 디자인 개선 — 이미지 카드 + 골드 호버 + 라운드 카드 (`7aa062f`)
- **feat:** 카테고리 배너 이미지 + 텍스트 오버레이 통합 — 블랙 바/타이틀/구분선 제거, `<picture>` PC/모바일 반응형 (`6a8cb1e`)
- **style:** PC 1300px 고정 + 태블릿~모바일 100% + 카테고리 탭 블랙 칩 형태 (`73fb61d`)
- **refactor:** 상품 진열 CSS를 공통 코드(Header Code)용 스크립트로 전환 — 위젯 ID 배열(`IDS`)로 일괄 관리, 페이지별 스타일 코드위젯 불필요 (`22eb8c2`)
- **docs:** 아임웹 공통 코드 백업 파일 생성 — `projects/common_code/` 폴더 신설 (`83c4a2d`)

### 2026-02-03 (A/S 접수)
- **style:** AS접수 페이지 가독성 개선 — 섹션 구분선 5개 삽입 + 간격 조정 (`3324f94`)
- **fix:** PC에서 수거요청일 달력 두번 클릭해야 열리는 문제 수정 — `openCalendar()` 중복 호출 가드 (`06488fa`)
- **fix:** 달력 모달 열기/닫기 잔상 제거 — overlay/sheet transition 0.2s 동기화, `removeAttribute('style')` 타이밍 개선 (`52b084a`)
- **fix:** `iframe_form.html`에 `MAMORU_REQUEST_VIEWPORT` 핸들러 추가 — iframe 내 모달 중앙 배치 지원 (`52b084a`)
  > **아임웹 /53 코드위젯도 동일 내용으로 교체 필요**

---

## 참고
- 상세 작업 규칙/역할/출력 포맷은 `CLAUDE.md` 및 `.claude/ADDENDUM_*`를 따릅니다.
- 브랜드/기획 전체본(참고): `docs/reference/BRAND_GUIDE_FULL.md`