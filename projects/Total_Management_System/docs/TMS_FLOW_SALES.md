# 판매관리 프로세스 흐름도
> 최종 업데이트: 2026-03-21

---

## 1. 비즈니스 프로세스 흐름

### 오프라인 판매 등록
```
(매장 방문 고객)
  → (관리자) TMS /sales/new 판매 입력
    → 카테고리별 제품 선택 (BL/TH/LO/SL)
    → 수량/할인/결제방법 입력
    → 고객명 + 전화번호 입력
    → "판매 등록" → sale_number 자동 채번 (OS-YYYYMMDD-NNN)
  → /sales/[id] 상세 페이지
```

### 결제 유형
```
card    — 카드 결제
cash    — 현금 결제
transfer — 계좌이체
mixed   — 복합 결제
```

### 결제 상태
```
paid    — 결제완료 (paid_amount = total_amount - discount_amount)
unpaid  — 미결제
partial — 부분결제
```

---

## 2. 시스템 연동 흐름

```
[관리자]
  │
  ▼
[TMS /sales/new] ──POST──→ [/api/sales]
                               │
                               ├──→ [Supabase DB] offline_sales + offline_sale_items
                               ├──→ [TMS 재고 차감] products.stock_quantity -N
                               ├──→ [아임웹 재고 동기화] PATCH /v2/shop/products/{no} (자동)
                               └──→ [미수금 반영] 미결제 시 customers.outstanding_balance 업데이트
                               │
                               ▼
                          [/sales/[id]]
                          판매 상세 페이지
```

---

## 3. 구현 완료 ✅

### API Routes (1개)
| 엔드포인트 | 메서드 | 기능 |
|------------|--------|------|
| `/api/sales` | GET | 목록 조회 (검색/페이징) |
| `/api/sales` | POST | 판매 등록 (sale + items) |

### 페이지 (3개)
- `/sales` — 판매 목록 (결제상태 배지, 검색, 페이징)
- `/sales/new` — 판매 입력 (카드형 제품 선택 + 장바구니 + 결제)
- `/sales/[id]` — 판매 상세

### DB 테이블
- `offline_sales` — 판매 메인 (sale_number, 고객, 결제)
- `offline_sale_items` — 판매 상품 (제품, 수량, 단가)
- `customers` — 고객 정보

---

## 4-A. 추가 구현 완료 ✅ (2026-03-01)

| 항목 | 구현 내용 |
|------|-----------|
| 고객 자동완성 | GET /api/customers/search — 이름/전화번호 ILIKE 검색, CustomerAutocomplete 공유 컴포넌트 |
| 고객 신규등록 | POST /api/customers — customers INSERT |
| 계약서 고객 자동완성 | contracts/new에도 CustomerAutocomplete 적용 (email/address 확장 필드) |

## 4-C. 추가 구현 완료 ✅ (2026-03-21)

| 항목 | 구현 내용 |
|------|-----------|
| 재고 연동 | 판매 시 TMS 재고 자동 차감 + 아임웹 재고 동기화 |
| COGS/마진 | 회계에서 제품별 매출원가·이익·이익률 분석 |
| 매입처 드롭다운 | SupplierSelect 컴포넌트 — 제품/발주 폼 |
| 재고 수동 조정 | 파손/실사 보정 모달 + stock_adjustments 이력 |
| 미수금 자동 반영 | 미결제 판매 시 customers.outstanding_balance 업데이트 |
| 마스터-디테일 | 제품 목록 클릭 → 우측 패널 상세 표시 |

## 4-D. 미완료 ❌

| 항목 | 의존성 | 우선순위 |
|------|--------|----------|
| 판매 수정/삭제 기능 | 없음 | 낮음 |

---

## 5. 핵심 파일 맵

### TMS API
| 파일 | 설명 |
|------|------|
| `app/src/app/api/sales/route.ts` | GET/POST 판매 |

### TMS UI
| 파일 | 설명 |
|------|------|
| `app/src/app/(dashboard)/sales/page.tsx` | 판매 목록 |
| `app/src/app/(dashboard)/sales/new/page.tsx` | 판매 입력 (카드형 UI) |
| `app/src/app/(dashboard)/sales/[id]/page.tsx` | 판매 상세 |

### TMS Lib
| 파일 | 설명 |
|------|------|
| `app/src/hooks/use-sales.ts` | React Query 훅 |

### DB 마이그레이션
| 파일 | 설명 |
|------|------|
| `app/supabase/migrations/007_offline_sales.sql` | offline_sales + offline_sale_items |
