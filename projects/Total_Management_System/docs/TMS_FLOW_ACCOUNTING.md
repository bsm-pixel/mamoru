# 회계 모듈 프로세스 흐름도

> 최종 수정: 2026-05-12 — **매출 3분할 (제품 B2C/B2B + 복원수리 A/B/C) + 회계 리포트 RS 집계 정확화** (코드 완료, push 대기)

## 2026-05-12 — 매출 3분할 (2-D 회계 리포트)

`api/reports/summary/route.ts` 개편:
- **제품 매출** = offline_sales (total−discount − 그 주문 RS_total) + **deliveries** (total−discount − 납품 RS_total). B2C(소매·온라인) / B2B(딜러·아카데미 offline + 납품 전체) 분리. ← 기존엔 deliveries 가 회계 리포트에서 통째로 누락돼 있었음.
- **복원수리 매출** = A(repairs, paid_at 기준 실제 금액 service_cost+shipping_fee) + B(offline_sale_items category='RS' total_price) + C(delivery_items category='RS' total_price). ← 기존엔 A(repairs)만 집계.
- **by_product / margin** — RS 항목 제외(RS는 제품 아님 → 제품 랭킹·원가 오염 제거), 납품 제품도 합산.
- **total_revenue** = 제품 매출 + 복원수리 매출 (offline RS 중복 계상 제거).
- offline_sales 조회에 `cancelled_at IS NULL AND returned_at IS NULL` 필터 추가 (취소·반품 매출이 회계에 잡히던 버그 fix).
- 응답: `product_sales{total,b2c,b2b,b2b_offline,b2b_delivery,offline_count,delivery_count}`, `repair_sales{total,a_intake,b_offline_rs,c_delivery_rs,a_count,b_count,c_count,bc_qty,service_cost_total,shipping_fee_total,count}`, `details.deliveries` 추가. `sales` 객체(offline 전체, RS 포함)는 호환용 유지.
- `reports/page.tsx` 매출 카드 = 전체(제품·복원수리 2분할 + 하위 내역) / 제품(B2C·B2B오프라인·B2B납품) / 복원수리(접수·판매RS·납품RS) 탭별 표시.
- ⚠️ 한계: 복원수리 A채널만 paid_at 기준이고 B/C는 sale_date/delivery_date 기준 — 회계 정밀화 필요시 추후 날짜 기준 통일.

## 2026-04-30 정리 (075)

| 변경 | 내용 |
|------|------|
| **복원수리 매출 정의 통일 (옵션 A)** | A채널(접수)+B채널(판매 RS)+C채널(납품 RS) 모두 **발생 기준** (sale_date/created_at/delivery_date) — 미입금도 매출로 카운트, 취소만 제외 |
| **경비 카테고리 동적화** | 설정 → 회계 → 경비 카테고리 추가/수정 → 즉시 /expenses 화면 반영 (이전엔 hard-coded) |
| **일정 재요청 카운트 fix** | 대시보드 needAction에서 pending_admin 잘못 포함 제거 (신규 상담이 재요청으로 카운트되던 버그) |
| **invalidate 풀 연동** | `lib/query/invalidate-keys.ts` 신설 → 모든 sale/repair mutation 후 대시보드 매출/통계 즉각 갱신 |

---

## 1. 회계 리포트 (`/reports`)

### 데이터 소스 (075 발생 기준 적용)
| 소스 | 테이블 | 조건 |
|------|--------|------|
| 상품 매출 | `offline_sales` | sale_date 범위, 취소 제외 |
| 복원수리 매출 A (접수) | `repairs` | **created_at 범위, 취소 제외** (paid_at 무관 — 075) |
| 복원수리 매출 B (판매 RS) | `offline_sale_items` | category='RS', sale_date 범위, 0원 제외, 취소 제외 |
| 복원수리 매출 C (납품 RS) | `delivery_items` | category='RS', delivery_date 범위, status≥confirmed, 취소 제외 |
| 매입 | `purchase_orders` | order_date 범위, 상태 filtered |
| 경비 | `expenses` | expense_date 범위 |

### 요약 카드
- **매출**: 상품 매출 + 복원수리 합산 (탭: 전체/상품/복원수리)
- **매입**: 발주 합계
- **부가세**: 매출세액 - 매입세액 = 납부세액
- **손익**: 매출 - 원가(COGS) - 경비 = 영업이익

### 파일
| 파일 | 역할 |
|------|------|
| `app/api/reports/summary/route.ts` | 기간별 집계 API (매출+매입+VAT+COGS+경비+복원수리+에이징) |
| `app/api/reports/export/route.ts` | 엑셀 다운로드 |
| `app/(dashboard)/reports/page.tsx` | 회계 대시보드 UI |
| `hooks/use-reports.ts` | 데이터 fetching + 타입 |

---

## 2. 경비 관리 (`/expenses`)

### 흐름
```
경비 등록 (날짜+카테고리+금액+메모) → expenses 테이블
                                        ↓
                                  회계 리포트 → 손익계산서 경비 합산
```

### 고정 경비
```
고정 경비 등록 (임대료/인건비 등) → recurring_expenses 테이블
                                        ↓
                              "이번달 일괄 등록" 클릭
                                        ↓
                              expenses 테이블에 [고정] 태그로 생성 (매월 1일)
```

### 파일
| 파일 | 역할 |
|------|------|
| `app/api/expenses/route.ts` | CRUD (GET/POST/DELETE) |
| `app/api/expenses/recurring/route.ts` | 고정 경비 관리 + 일괄 생성 |
| `app/(dashboard)/expenses/page.tsx` | 경비 페이지 (등록+목록+고정경비) |

### DB
- `expenses`: id, expense_date, category, amount, memo, created_by
- `recurring_expenses`: id, category, amount, memo, is_active

---

## 3. 입출금 관리 (`/cashflow`)

### 흐름
```
입금 등록 (매출입금/기타) → cash_transactions (type=income)
출금 등록 (매입결제/경비) → cash_transactions (type=expense)
                              ↓
                        요약: 입금 합계 / 출금 합계 / 순이익
```

### 파일
| 파일 | 역할 |
|------|------|
| `app/api/cashflow/route.ts` | CRUD (GET/POST/DELETE) |
| `app/(dashboard)/cashflow/page.tsx` | 입출금 페이지 |

### DB
- `cash_transactions`: id, transaction_date, type(income/expense), category, amount, memo

---

## 4. 세금계산서 (`/tax-invoices`)

### 흐름
```
세금계산서 등록 (매출/매입 + 거래처 + 공급가 + 세액)
                              ↓
                        분기별 요약: 매출세액 / 매입세액 / 납부세액
                              ↓
                        부가세 신고 시 참조
```

### 자동 계산
- 공급가 입력 → 세액 = 공급가 × 10% 자동

### 파일
| 파일 | 역할 |
|------|------|
| `app/api/tax-invoices/route.ts` | CRUD (GET/POST/DELETE) |
| `app/(dashboard)/tax-invoices/page.tsx` | 세금계산서 페이지 |

### DB
- `tax_invoices`: id, invoice_type(sales/purchase), issue_date, counterparty_name, counterparty_biz_no, supply_amount, tax_amount, total_amount

---

## 5. 매출채권 에이징

### 회계 리포트 내 표시
```
고객별 미수금 (outstanding_balance > 0)
  ↓
가장 오래된 미결제 판매일 조회 (offline_sales)
  ↓
경과일 계산 → 구간 분류
  30일 이내 (양호/초록)
  30~60일 (주의/노랑)
  60~90일 (경고/주황)
  90일+ (위험/빨강)
```

---

## 6. 손익계산서 구조

```
매출 (상품 + 복원수리)
  - 매출원가 (COGS = 제품별 매입가 × 수량)
  ─────────────────────
  = 매출총이익
  - 경비 (expenses 테이블 기간 합계)
  ─────────────────────
  = 영업이익
  영업이익률 = 영업이익 / 매출 × 100
```

---

## 7. 채널별 매출 분류 (이번달 매출이 어디서 합산되나)

### 매출 발생 → 합산 흐름

```
┌─────────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐
│   판매 (offline_    │    │   복원수리          │    │   B2B 납품          │
│   sales)            │    │   (repairs)         │    │   (deliveries)      │
│                     │    │                     │    │                     │
│ • 매장 결제          │    │ • 입금 확정 시점     │    │ • 납품 확정 시점      │
│ • 톡상담 결제        │    │   매출 인식         │    │   매출 인식          │
│ • 온라인 결제        │    │ • paid_at 기준      │    │ • paid_at 기준      │
│ • B2B 즉시 납품      │    │                     │    │                     │
└──────────┬──────────┘    └──────────┬──────────┘    └──────────┬──────────┘
           │                          │                          │
           ▼                          ▼                          ▼
        ┌──────────────────────────────────────────────────────────────┐
        │  대시보드/리포트 매출 합산 (revenue_basis 설정 기준)              │
        │                                                              │
        │  • sale_date (판매일)  ← 기본값                                │
        │  • payment_date (결제일)                                      │
        │  • ship_date (출고일)                                         │
        └──────────────────────────────────────────────────────────────┘
                                      │
                                      ▼
                          ┌────────────────────────┐
                          │  손익계산서 / 대시보드   │
                          │  • 이번달 매출           │
                          │  • 카테고리별            │
                          │  • 채널별                │
                          └────────────────────────┘
```

### 어디서 합산이 일어나는가

| 매출 종류 | DB 위치 | 인식 시점 | 페이지 |
|-----------|---------|----------|--------|
| **상품 판매** | `offline_sales` | sale_date / paid_at / shipped_at | /sales, /reports |
| **복원수리** | `repairs.paid_at` | 입금 확정일 | /repairs, /reports |
| **B2B 납품** | `deliveries.paid_at` | 정산 완료일 | /deliveries, /reports |
| **별도 입금** | `cash_transactions` | transaction_date | /cashflow |

### 입출금(현금흐름) vs 매출의 차이

- **매출**: 위 3개 테이블이 본 출처 → 손익계산서/대시보드에 자동 반영
- **입출금(cash_transactions)**: 별도 입금/출금 기록 (보증금, 환불, 투자금 등 이벤트성 자금 흐름)
  - 입금 카테고리: 매출입금 / 기타입금 / 환불수령 / 투자금
  - 출금 카테고리: 매입결제 / 경비 / 세금 / 인건비 / 임대료 / 기타출금
  - **카테고리는 설정 → 회계에서 자유롭게 추가/삭제 가능**

> ⚠️ "매출입금" 카테고리는 별도 입금 기록용이며, 손익계산서의 매출 합산과 **이중 계상되지 않음** (cashflow는 운영 자금 흐름 추적용).

### 카테고리 설정 위치

설정 → 회계 탭 → "입출금 — 입금 카테고리" / "입출금 — 출금 카테고리"
- 사장님이 직접 추가/삭제 가능
- 미저장 시 기본값 자동 사용
- 기존에 등록된 입출금 레코드는 카테고리 삭제해도 그대로 유지 (orphan 안전)
