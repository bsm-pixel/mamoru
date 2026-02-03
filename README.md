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

## 4. 구현 가이드라인(요약)
- **Performance:** 로고는 SVG, 애니메이션은 Lottie 지향. 이미지 배너 남발 대신 코드/스타일 위젯 우선.
- **Data-Driven:** 콘텐츠(Text/Image)와 로직(Code) 분리. 시트 변경만으로 운영 가능하게 설계.
- **Responsive:** Mobile-first. PC에서는 max-width + 중앙정렬로 과도한 확장 방지.
- **Micro-copy:** 기능 나열/가격 강조 대신, 동행/책임/안심이 느껴지는 문장 사용.

---

## 참고
- 상세 작업 규칙/역할/출력 포맷은 `CLAUDE.md` 및 `.claude/ADDENDUM_*`를 따릅니다.
- 브랜드/기획 전체본(참고): `docs/reference/BRAND_GUIDE_FULL.md`