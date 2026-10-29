# 회계 모듈 프로세스 흐름도

> 최종 수정: 2026-10-29 — 영수증·세금계산서 메일 → 경비 확인 대기 + 고정/사용량 구분 (마이그 166)
> 이전: 2026-10-12 — 할인 반영 일원화(할인 후 최종금액만 매출) + 직접 입력 품목 분류 + 납품 할인 이중차감 fix(마이그 160)

## 2026-10-29 — 메일 영수증 → 경비 (마이그 166)

배경(사장님): 경비를 손으로 넣다 보니 6월 이후 입력이 멈춤. 영수증 메일은 bsm@mamoru.kr 로 모인다
(mamoruscissors@gmail.com · bsm19911@gmail.com 은 그쪽에서 자동 전달 — 메일함 하나만 읽는다).

```
매일 08:00 KST 크론 /api/cron/expense-mail  (또는 경비 화면 [메일 다시 확인] · 처음 1회 [지난 4개월 가져오기])
  → Gmail 검색(영수증·결제 내역·세금계산서·거래명세서…, 최근 3일) — 구글 연결 토큰, gmail.readonly
  → lib/expenses/mail-parse.ts parseExpenseMail (규칙만, 유료 판독 X)
       KG파이낸셜·NHN KCP·키움페이 카드결제 / Google Play / Anthropic($) / 팝빌·유니포스트 세금계산서 / 구글 워크스페이스·솔라피 청구서(금액=PDF → 빈칸)
       공지·검수·고객 주문 알림·[MAMORU] 자체 알림은 거른다
  → expense_inbox (status=pending, gmail_id UNIQUE — 같은 메일 1번만)
       거래처 기억 expense_vendor_rules: 분류·고정/사용량 자동, auto_ignore 면 바로 ignored
  → 경비 화면 「메일에서 찾은 영수증」: 금액·날짜·분류·고정/사용량 확인 → [확인]
       → expenses 1건(kind·vendor·source_ref='mail:<id>') + 거래처 기억 저장
       [무시] / [늘 무시](그 거래처 다음부터 자동 무시)
```

- **회계 숫자는 사람이 확인한 것만**: pending 은 expenses 에 없음 → 리포트 손익·경비 합계에 안 잡힌다.
- 중복 의심: 같은 금액 경비가 ±7일 안에 있으면 행에 안내(카드 영수증 + 세금계산서 이중 등).
- 외화(USD) 영수증은 원금만 기록, 원화 청구액은 [확인] 때 입력. 메모 끝에 `(USD 100)` 남김.
- `expenses.kind` fixed=고정(구독·임대) / variable=사용량(택배·부자재). 경비 화면 요약에 합계, 옛 데이터는 미분류.
  고정 경비 [이번달 일괄 등록]은 kind=fixed.
- 메일로 안 오는 지출(현장 카드·계좌이체)은 아직 수기 — 다음 단계 후보: 출금 알림(은행 입금 자동확인과 같은 MacroDroid 경로).

## 2026-10-12 — 할인 반영 일원화 + 직접 입력 품목 분류 (마이그 160)

사장님 지적: "이번달 매출이 560만원대였는데 550만원으로 줄었다."
→ **총액은 정상**이었다. 라이브 `get_hub_stats` 재집계 결과 5,498,800원이 맞는 값(원시 데이터와 대조 일치). 560만은 금액/할인 수정 전 화면으로 추정되며, `offline_sales` 에는 변경 이력 테이블이 없어 **무엇이 언제 바뀌었는지는 증명 불가**(향후 필요하면 감사 테이블 신설 건으로 남김).

다만 파고드는 과정에서 **할인·분류 쪽 실제 문제 3건**이 나와 한 번에 고쳤다.

### ① 직접 입력 품목에 분류가 없어 복원수리가 제품 매출로 잡힘 (근본)
- 원인: 판매 등록의 `addCustomProduct` 가 `category: undefined` 로 저장 → `offline_sale_items.category` 가 NULL. 복원수리 매출 B채널은 `category='RS'` 로만 걸러지므로, 손으로 적어 넣은 복원수리는 **제품 매출**에 들어갔다.
- fix: `/sales/new` + `FullEditSaleModal`(sale-detail-panel) 두 입력 경로 모두에 **[제품] / [복원수리] 토글** 신설. 선택에 따라 안내 문구도 "제품 매출로 집계됩니다 / 복원수리 매출로 집계됩니다" 로 바뀐다.
- 기존 데이터 보정: NULL 이던 복원수리 품목 6건을 `category='RS'` 로 고침.
  - **총매출 불변** 5,498,800원 / 분해만 교정: B2C 4,756,000 → 4,711,000, 복원수리 463,000 → 508,000.

### ② 거래내역서 합계가 할인 전 금액 (`/reports/transaction`)
- 원인: 합계를 `total_amount` 로만 더해 할인이 빠지지 않았다(거래처에 보내는 문서인데 실수령액과 달랐음).
- fix: `netOf(r) = total_amount − discount_amount` 로 행·소계·총계 통일. 할인 있는 행엔 `할인 -N원` 보조줄, 총계엔 `할인 -N원 반영` 표기.
- 공급가액·부가세: DB 저장값이 할인 전 기준이라 합계와 안 맞던 문제 → `vatSplit()` 로 **저장된 공급가:부가세 비율을 유지한 채 net 기준으로 재배분**(DB 값은 건드리지 않음, 화면 계산만). 검증: 620,000 = 563,636 + 56,364.

### ③ 납품 할인 이중차감 (RPC 088 → 마이그 **160**)
- 원인: `get_hub_stats` 가 `deliveries` 에 `SUM(total_amount − discount_amount)` 를 적용. 그런데 금액 SSOT(`lib/sales/amounts.ts`)대로 **`deliveries.total_amount` 는 이미 NET** — 할인이 두 번 빠졌다.
- fix: `160_hub_stats_delivery_discount_fix.sql` — 088 전문을 그대로 재생성하되 **두 줄만** `SUM(total_amount)` 로 교체(`deliveries` 직접 집계 + `dl.` 별칭 경로). `offline_sales` 쪽 할인 차감은 **의도된 것이므로 유지**(그쪽 total_amount 는 GROSS).
- 실제 영향 건: `DL-20260914-001` (9월, 20,000원). 10월분엔 할인 납품이 없어 이번달 숫자는 변동 없음.

### 회계 원칙 (사장님 판단 요청에 대한 답 — 확정)
**할인 후 최종금액만 매출로 인식한다.** 화면·리포트·RPC 전부 이 기준.
테이블별 저장 방식이 달라 계산식도 다르다 — 반드시 SSOT(`lib/sales/amounts.ts`)를 쓸 것:

| 테이블 | `total_amount` 의미 | 매출 계산식 |
|--------|--------------------|------------|
| `offline_sales` | **GROSS** (할인 전) | `total_amount − discount_amount` |
| `deliveries` | **NET** (할인 반영됨) | `total_amount` (재차감 금지) |

### 남은 것 (별도 결정 필요)
- **복원수리 날짜 기준 불일치**: 대시보드는 A채널 `created_at`, 회계 리포트는 `paid_at` — 어느 쪽으로 통일할지 사장님 결정 사항.
- **결제수단별 매출·부가세의 할인 반영**: 세무 처리와 맞물려 세무사 확인 후 진행.
- RPC 088 `returned_at` 비대칭(소급 vs 반품시점), `useDeliveryStats` 상태 필터 부재(초안 납품 포함), `api/reports/summary` 기본 from/to UTC — 이번 범위 밖, 기록만.

## 2026-08-01 — 취소/반품 제외 일관화 (A+B+C) + EVENT접수취소 활성판매 차단
- **원칙 재확인**: 취소(cancelled)=회계 완전 제외 / 반품(returned)=매출 리포트는 순매출(원 판매월에서 소급 제거, `returned_at IS NULL`). 현금흐름은 반품월 환불 반영(대칭). ※ "반품을 매출 리포트에도 반품월 −반영"으로 바꾸는 건(D) 별도 회계정책 결정 — 보류.
- **B(버그) fix**: 엑셀 내보내기 `api/reports/export/route.ts` 매출·마진 2곳이 취소·반품 **무필터**였음 → `cancelled_at/returned_at IS NULL` 추가.
- **C(보강)**: `use-dashboard-stats.ts` 대시보드 fallback·RPC077 경로·복원수리 RS(월/주) 6곳이 `returned_at` 필터 누락 → 추가(취소는 이미 있었음). RPC088 메인합은 원래 정상.
- **A(재발방지)**: EVENT/재고판매 접수 취소(`api/events/[id]` action='cancel') 시, 연결 판매(sale_id)가 **아직 활성이면 409로 차단**하고 "판매관리에서 판매 먼저 취소/반품" 안내. → 접수만 취소해 판매가 유령으로 남아 매출에 잡히던 문제 예방. (판매가 이미 취소/반품됐으면 접수 정리는 허용)
## 2026-08-01 (2차) — D 반품시점 반영(reports/summary) + E 세금계산서 경고
- **D 정책 전환**: 반품 매출 = **소급 제거 → 반품시점 −반영**. 판매는 `sale_date`(판매월)에 인식, 반품은 `returned_at`(반품월)에 −차감. 판매·반품 동월이면 net 0.
  - **reports/summary(품목별 매출) 적용 완료**: sales 쿼리에서 `returned_at IS NULL` 제거(판매월 인식) + `retSales`(returned_at 기준) 별도 −차감. by_product/마진은 반품 제품항목을 음수로 넣어 자동 net. saleSummary.total·dailyProduct·dailyRepairs·offlineRs 모두 −반영. 딜리버리/온라인주문은 반품개념 없어 무변경.
  - **useSalesStats(판매 요약 카드) 적용 완료(2차)**: week/month·customer/partner·RS·b2b 랭킹 전부 판매월 +/반품월 − (returned 필터 제거 + *Returns 쿼리 −차감). count 는 판매월 기준 유지. 반품 검증(reports/summary) 통과 후 동일 패턴 이식.
  - ⚠️ **대시보드는 아직 소급**: `use-dashboard-stats` + **RPC 088**(get_hub_stats, 265줄 Postgres 함수, 대시보드 메인 경로)은 1차 C 의 `returned_at IS NULL`(소급) 유지. RPC 손수 재작성은 DB 검증 불가라 위험 → **보류**. 영향: 대시보드 '이번달 매출' glance 가 **cross-month 반품일 때만** reports/카드와 다름(동월 반품·취소는 동일). 반품 드물어 실무 영향 작음. 원하면 RPC 마이그를 별도 작성(사장님 Supabase 실행+검증).
  - **검증**: 앱 계층 tsc/build 통과. ⚠️ DB 숫자는 미검증 — 실제 반품 1건으로 "판매월 +, 반품월 −, 전체기간 net 0" 눈으로 확인 필요.
- **E 세금계산서 경고(자동취소 X)**: `tax-invoices` GET 이 연결 판매(sale_id) 취소/반품 시 `linked_sale_voided` 플래그 → 목록에 "⚠️ 수정세금계산서 필요" 배지. **원본 자동삭제 안 함**(세법상 반품/취소는 수정세금계산서 별도 발행 = 수기 절차).

## 2026-06-18 — 일별 추이 차트 탭별 분리
- 기존: 일별 차트가 탭 무관하게 `daily.sales`(offline_sales 제품 일별, **납품 미포함·RS 혼입**)만 그림 → 복원수리 탭에서도 제품 일별만 보여 B2B 복원수리(납품)가 차트에 안 나옴.
- 변경: route `daily.product`(제품 = 판매−RS + 납품제품) 신설, 기존 `daily.repairs` 활용. 페이지가 탭별로 시리즈 선택 — **전체=제품+복원수리 / 제품=제품만 / 복원수리=복원수리만(접수 paid_at + 판매RS + 납품RS, 납품 포함)**. 복원수리 탭은 매입(제품 원가성) 숨김.
- 참고: 회계 **총액 집계는 정상**(납품 RS는 status≥confirmed·취소제외로 이미 잡힘). 차트만 누락이었음.

## 2026-05-13 후속 (배포됨)
- 회계 "이번 달" 프리셋이 `toISOString().slice(0,10)`(UTC)라 5/1 아닌 4/30부터 잡히던 버그 → `toLocalDateString` 으로 fix.
- 매출 카드 "제품 판매" 탭: "B2C 제품 (소매·온라인)" / "B2B 제품 (딜러·아카데미·납품)" 두 줄로 통합 (이전 "B2B 딜러·아카데미 0원" 군더더기 줄 제거, >0 일 때만 하위 분해 표시). 탭 라벨 '상품 판매'→'제품 판매'.
- 배송비: 복원수리 거래의 "배송비 3,000원" 항목은 복원수리 매출 금액엔 포함(접수시스템 `shipping_fee` 와 통일)하되, **"자루 수"(`repair_sales.bc_qty`) 카운트에선 제외** (`product_name <> '배송비'`). RPC 080 + 클라이언트 훅 + 회계 route 전부 반영.

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
