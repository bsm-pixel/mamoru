# TMS 파일 수정 가이드

> 문구/명칭/UI를 직접 수정할 때 어디를 수정하면 되는지 안내
> 기준 경로: `projects/Total_Management_System/app/src/`

---

## 1. 사이드바 · 네비게이션 메뉴 명칭

| 수정 대상 | 파일 |
|-----------|------|
| PC 사이드바 메뉴 이름/아이콘/순서/그룹명 | `lib/utils/constants.ts` → `NAV_GROUPS` |
| 모바일 하단 탭 바 (고정 4개 + 더보기) | `components/layout/mobile-nav.tsx` |
| 사이드바 아이콘 추가 시 import | `components/layout/sidebar.tsx` → `iconMap` |

**예시:** "판매관리" → "판매"로 변경 → `constants.ts`에서 `label: '판매관리'`를 `label: '판매'`로 수정

---

## 2. 각 페이지별 수정 위치

### 대시보드
| 페이지 | 파일 |
|--------|------|
| 허브 대시보드 (메인) | `app/(dashboard)/dashboard/page.tsx` |
| 주문 대시보드 | `app/(dashboard)/orders/dashboard/page.tsx` |
| 상담 대시보드 | `app/(dashboard)/consultations/dashboard/page.tsx` |
| 복원수리 대시보드 | `app/(dashboard)/repairs/dashboard/page.tsx` |

### 영업 · 계약
| 페이지 | 파일 |
|--------|------|
| 고객 목록 | `app/(dashboard)/customers/page.tsx` |
| 고객 상세 | `app/(dashboard)/customers/[id]/page.tsx` |
| 계약서 목록 | `app/(dashboard)/contracts/page.tsx` |
| 계약서 작성 (전자문서) | `app/(dashboard)/contracts/new/page.tsx` |
| 계약서 상세 | `app/(dashboard)/contracts/[id]/page.tsx` |
| 판매 목록 | `app/(dashboard)/sales/page.tsx` |
| 판매 입력 | `app/(dashboard)/sales/new/page.tsx` |
| 판매 상세 | `app/(dashboard)/sales/[id]/page.tsx` |

### CS · 수리
| 페이지 | 파일 |
|--------|------|
| 상담 목록 | `app/(dashboard)/consultations/page.tsx` |
| 상담 상세 | `app/(dashboard)/consultations/[id]/page.tsx` |
| 복원수리 목록 | `app/(dashboard)/repairs/page.tsx` |
| 복원수리 상세 | `app/(dashboard)/repairs/[id]/page.tsx` |

### 물류 · 재고
| 페이지 | 파일 |
|--------|------|
| 주문 목록 | `app/(dashboard)/orders/page.tsx` |
| 주문 상세 | `app/(dashboard)/orders/[id]/page.tsx` |
| 매입 목록 | `app/(dashboard)/purchasing/page.tsx` |
| 매입 작성 | `app/(dashboard)/purchasing/new/page.tsx` |
| 매입 상세 | `app/(dashboard)/purchasing/[id]/page.tsx` |
| 재고 현황 | `app/(dashboard)/inventory/page.tsx` |

### 상품
| 페이지 | 파일 |
|--------|------|
| 제품 목록 | `app/(dashboard)/products/page.tsx` |
| 제품 등록 | `app/(dashboard)/products/new/page.tsx` |
| 제품 상세 | `app/(dashboard)/products/[id]/page.tsx` |
| 시리얼 관리 | `app/(dashboard)/products/[id]/serials/page.tsx` |

### 정산
| 페이지 | 파일 |
|--------|------|
| 회계 리포트 | `app/(dashboard)/reports/page.tsx` |
| 거래내역서 (인쇄) | `app/(dashboard)/reports/transaction/page.tsx` |

### 시스템
| 페이지 | 파일 |
|--------|------|
| 설정 | `app/(dashboard)/settings/page.tsx` |
| 로그인 | `app/(dashboard)/login/page.tsx` |

---

## 3. 계약서 전자문서 법적 문구

| 수정 대상 | 파일 | 위치 |
|-----------|------|------|
| "마모루는 도움을 드리는 회사입니다" 본문 | `app/(dashboard)/contracts/new/page.tsx` | 상단 `LEGAL_NOTICE` 변수 |
| 유의사항 (교환/반품 조건) | `app/(dashboard)/contracts/new/page.tsx` | 상단 `CAUTION_NOTICE` 변수 |
| 입금 계좌 정보 | `app/(dashboard)/contracts/new/page.tsx` | 하단 "입금 계좌" 텍스트 |

---

## 4. 공통 컴포넌트

| 컴포넌트 | 파일 | 용도 |
|----------|------|------|
| 사이드바 (PC) | `components/layout/sidebar.tsx` | 좌측 메뉴 |
| 모바일 네비 | `components/layout/mobile-nav.tsx` | 하단 탭 바 + 더보기 |
| 상단 바 | `components/layout/topbar.tsx` | 페이지 제목 + 버튼 |
| 고객 자동완성 | `components/shared/customer-autocomplete.tsx` | 고객 검색/선택/신규등록 |
| 서명 캔버스 | `components/contracts/signature-canvas.tsx` | 터치/마우스 서명 |
| 제품 선택 모달 | `components/contracts/product-picker-modal.tsx` | 계약서 제품 선택 |
| 시리얼 선택 | `components/sales/serial-picker.tsx` | 판매 시 시리얼 연결 |
| 버튼/카드/뱃지 등 | `components/ui/` 폴더 | UI 기본 요소 |

---

## 5. 상태 라벨 · 한글 매핑 (자주 수정하는 것들)

| 대상 | 파일 | 변수명 |
|------|------|--------|
| 주문 상태 라벨 | `lib/utils/format.ts` | `STATUS_LABEL` / `STATUS_COLOR` |
| 상담 상태 라벨 | 각 페이지 내 `STATUS_LABEL` 객체 |
| 복원수리 상태 라벨 | 각 페이지 내 `STATUS_LABEL` 객체 |
| 결제방식 라벨 (카드/현금/이체) | 각 페이지 내 `PAYMENT_LABEL` 객체 |
| 택배사 코드/이름 | `lib/utils/constants.ts` → `COURIER_CODES` |
| 금액 포맷 (원) | `lib/utils/format.ts` → `formatKRW()` |
| 날짜 포맷 | `lib/utils/format.ts` → `formatDate()` / `formatDateTime()` |

---

## 6. API (데이터 저장/조회 로직)

| 도메인 | 경로 |
|--------|------|
| 고객 | `app/api/customers/` |
| 판매 | `app/api/sales/route.ts` |
| 계약서 | `app/api/contracts/route.ts` |
| 제품 | `app/api/products/` |
| 매입 | `app/api/purchasing/` |
| 재고 | `app/api/inventory/route.ts` |
| 회계 | `app/api/reports/summary/` + `export/` |
| 상담 | `app/api/consultation/` |
| 복원수리 | `app/api/repair/` |
| 주문 | `app/api/orders/` |
| 시리얼 | `app/api/serials/` |

---

## 7. 수정 후 반영 방법

1. 파일 수정 후 저장
2. `git add` + `git commit` + `git push`
3. Vercel 자동 배포 (1~2분)

> 브라우저에서 직접 수정 불가 — VS Code 등 에디터에서 수정 필요
