# 회계 모듈 프로세스 흐름도

> 최종 수정: 2026-04-02

---

## 1. 회계 리포트 (`/reports`)

### 데이터 소스
| 소스 | 테이블 | 조건 |
|------|--------|------|
| 상품 매출 | `offline_sales` | sale_date 범위, 취소 제외 |
| 복원수리 매출 | `repairs` | paid_at 범위, 취소 제외 |
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
