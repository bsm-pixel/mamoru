# 판매관리 프로세스 흐름도
> 최종 업데이트: 2026-05-13 — 판매 입력 "복원수리" 모드 (3-A) + 거래처 복원수리 기본 단가 (2-E) + 배송비 토글 + RS 제품 숨김 (배포 완료)

## 2026-05-13 후속 (배포됨)
- `/sales/new` 제품 선택 목록에서 `category='RS'` 제품 숨김 + "RS 복원수리" 카테고리 필터 버튼 제거 — 복원수리는 상단 "복원수리" 모드로만 입력. (기존 RS 더미 제품 DB 레코드는 과거 판매 참조용으로 보존)
- `/sales/new` 복원수리 모드에 "배송비 3,000원 포함" 토글 — 켜면 `offline_sale_items` category='RS', product_name='배송비' 로 1행 추가 (복원수리 매출엔 포함, 자루 수엔 제외).
- `/deliveries` "+B2B수리" 의 추가항목·"+배송비 3,000원" 버튼에 `category='RS'` 부여 — 회계 리포트에서 복원수리 매출로 잡히게 (이전엔 category 없어 제품으로 오분류).

## 2026-05-12 — 판매 입력 "복원수리" 모드 (B채널) + 거래처 단가 (C채널 자동 채움)

- `/sales/new` 상단 모드 탭: **제품 판매 / 복원수리**.
  - 복원수리 모드: 좌측에 마모루 가위(기본 1만)/타사 가위(기본 2만) 각각 자루 수 + 단가 입력(조정 가능). 우측 장바구니 카드에 "복원수리 (마모루) ×N자루" / "복원수리 (타사) ×M자루" 라인 + 합계.
  - 저장: `offline_sale_items` category='RS', product_name 고정 "복원수리 (마모루)" / "복원수리 (타사)" → 매출 분류·대시보드 마모루/타사 집계와 호환(`includes('타사')`).
  - B2C 복원수리 입력 경로 완성. B2B 거래처 복원수리는 기존 `/deliveries` "+B2B수리" 모드 사용.
- `customers.default_repair_price` (마이그레이션 079): 거래처별 복원수리 자루당 단가. 고객 상세 화면(딜러/아카데미)에서 입력. `/deliveries` "+B2B수리" 모달에서 거래처 선택 시 단가 자동 채움(없으면 기본 8천원).

## 2026-04-30 (심야 +3) — /deliveries(B2B 거래/납품서)에도 catalog 자동 입력 적용

### 문제
사장님이 074 push 후 "B2B 거래 → 납품서 작성"에서 catalog 등록한 제품을 추가했는데 **모델명(A2-2522-PT)이 그대로** 표시됨. 074에서 /sales/new만 수정하고 /deliveries 누락.

### Fix
`/deliveries/page.tsx`에 동일 패턴 적용:
- `useCustomerCatalog(selectedCustomer?.id)` 호출
- `getPrice(p)` — 1순위 `catalog.unit_price`, 2순위 price_dealer/academy, 3순위 price
- 신규 `getDeliveryName(p)` — 1순위 `catalog.delivery_name`, 2순위 product.name
- `addProduct`에서 두 함수 모두 적용 → cart에 즉시 catalog 이름/가격 표시

### 회귀 안전
- catalog 미등록 제품 → 기존 fallback 그대로
- /sales/new 흐름 0 변경

---

## 2026-04-30 (심야 +2) — B2B 카테고리 동적화 + 거래처별 단가 (074)

### 핵심 변경
사장님 결정: B2B 카테고리(딜러/아카데미/학교/공기관 등)를 설정에서 자유롭게 추가/수정 + 각 거래처 catalog에 **납품가** 등록 → 판매 시 가격까지 자동.

### 마이그 074
- `system_settings`에 `b2b.categories` seed (dealer/academy 기본값, 사장님이 설정에서 추가)
- `customer_product_catalog`에 `unit_price` 컬럼 추가 (거래처별 맞춤 가격)

### B2B 카테고리 동적 관리
**설정 → 고객 관리 → "B2B 납품처 카테고리"**:
- 기본 (is_default=true): dealer, academy — 삭제 차단
- 사장님 추가 가능: school, public, hospital 등
- key (영문) + label (한글) + 아이콘 (8종) + 순서 + 활성 토글
- 추가 시 `/거래처` 페이지 탭 자동 노출

**/거래처 페이지 탭 동적**:
- B2B: 사장님 정의 (settings 기반)
- 매입처(supplier): system 고정 (catalog 흐름 다름)
- 색상: key 해시 기반 deterministic

### 가격 우선순위 (sales/new addToCart)
```
selectedCustomer 있을 때:
  1순위: customer_product_catalog.unit_price (074, 거래처별 맞춤가)
  2순위: product.price_groups[customerType].price (group 단가, 기존)
  3순위: product.price (소매가, 기본)
```

### 납품명 우선순위 (저장 시 product_name)
```
1순위: customer_product_catalog.delivery_name (073)
2순위: getProductDisplayName(product, customerType, priceGroups)
3순위: product.name
```

### catalog UI (3컬럼)
**납품명** / **납품가** / **특징** — 미입력 시 fallback 자동.

### 신규 카테고리 추가 시 사장님 흐름
1. 설정 → 고객 관리 → B2B 카테고리 → "추가" → key='school', label='학교', icon='School'
2. /거래처 → "학교" 탭 자동 노출
3. 거래처 등록 + 납품품목 catalog에 제품 + 납품가 + 납품명 등록
4. 판매 시 자동 적용 → 송장 출력에 박힘

### 회귀 안전
- customer_type DB 값 그대로 (사장님 추가 = 새 값 'school' 등)
- 29곳 hard-coded `customer_type === 'dealer'` 비교 그대로 (점진 helper 통합 별도 task)
- price_groups 그대로 (catalog.unit_price 미등록 시 fallback)
- supplier 흐름 0 변경 (system 고정)
- 070 link / 072 매칭 / 073 catalog (delivery_name) — 모두 영향 0

---

## 2026-04-30 (심야 +1) — B2B 납품처 카탈로그 + 납품명 자동 입력 (073)

### 변경 배경
사장님 결정: 매입처(supplier)에 있는 발주명 시스템을 B2B 납품처(dealer/academy)에 mirror로 복제. dealer A는 "MMR-150 BL", dealer B는 "마모루 150" 등 같은 제품을 납품처마다 다른 이름으로 부름 → 송장에 그 이름이 박혀서 출력.

### 구현
**마이그 073 — `customer_product_catalog` 테이블** (supplier_product_catalog와 별개):
- customer_id × product_id UNIQUE
- delivery_name (송장/명세서 출력용 납품명)
- features (규격/특이사항)

**API**: `/api/customers/[id]/catalog` GET/POST/PATCH/DELETE (supplier API mirror)
**Hook**: `useCustomerCatalog`, `useAddToCustomerCatalog`, `useUpdateCustomerCatalog`, `useRemoveFromCustomerCatalog`
**UI**: `/suppliers` 페이지 dealer/academy 탭에 "납품품목" 섹션 노출 (`CustomerCatalogSection` 컴포넌트)

### 자동 입력 흐름
```
/sales/new — 사장님이 customer 선택
  ↓ useCustomerCatalog(customer.id) fetch
  ↓ catalog에 등록된 product 매핑 (delivery_name)
저장 시 cart.map → offline_sale_items.product_name 결정:
  1순위: catalog.delivery_name (해당 customer의 등록된 납품명)
  2순위: price_groups display_name (customer_type별, 기존 fallback)
  3순위: product.name (기본)
  ↓
INSERT offline_sale_items
  ↓
모든 출력에 자동 반영:
  - 송장 (ALPS goodsName) — ship/route.ts가 product_name 그대로 사용
  - 거래명세서 (ReceiptModal)
  - 준비표 (PrepSheetModal)
```

→ 저장 시점 한 곳만 정확하면 모든 출력에 자동 반영. 송장 인쇄 코드 변경 X.

### 회귀 안전
- supplier 흐름 0 변경 (별도 테이블, 별도 API, 별도 UI)
- catalog 없는 dealer/academy → 기존 fallback (price_groups display_name → product.name)
- retail customer → catalog 미사용
- 이미 저장된 sale → product_name 보존 (catalog 변경되어도 거래 당시 이름 유지)

---

## 2026-04-30 (심야) — 고객 자동 매칭/생성과 판매 흐름 통합

상담/복원수리 접수 시 phone 기준 customer 자동 매칭/생성(`lib/customer/match-or-create.ts`)이 적용되어, **`/sales/new?from_consultation=<id>` 진입 시 SelectedCustomer가 자동 set**:
- 이전: from_consultation의 phone/이름 prefill (검색 단계 필요)
- 지금: consultation에 customer_id 채워져 있으니 SelectedCustomer 정식 연결 → customer_type 기반 가격 자동 적용

같은 고객 customer 상세 페이지에 상담/복원수리/판매 통합 표시.

상세: `docs/TMS_FLOW_CONSULTATION.md` "2026-04-30 심야" 섹션

---

## 2026-04-30 (밤) — 리뷰 진입점 단일화

상담관리 리뷰 연동이 전면 제거되어, **후기 약속/발송은 판매(sale) 또는 복원수리(repair) source에서만** 진행. 사장님이 sale 상세의 ReviewManagementCard에서 review_type을 자유롭게 선택(상담/복원수리/제품구매) → 솔라피 알림톡 정상 발송 → 후기 페이지가 sale_number를 fallback으로 매칭해 정상 동작.

### linkedConsultation 정보 칩 (070 link의 정보 가치 보존)
sale에 source_consultation_id가 있을 때 (출장/매장상담에서 시작된 거래), ReviewManagementCard 위에 작은 정보 칩 표시:
```
이 판매는 출장/매장상담 [CS-...]에서 시작됨
```
클릭 시 원본 상담 상세로 이동. mirror 박스(이전) 제거 — 정보 칩만.

### 양방향 가시화
- sale 상세 → 정보 칩으로 원본 상담 표시
- consultation 상세 → linkedSales 칩으로 후속 판매 표시 (변경 없음)

---

## 출장/매장상담 → 판매 link 인프라 (070, 2026-04-30 추가)

### 핵심 컬럼 — 마이그 070
`offline_sales.source_consultation_id UUID NULL REFERENCES consultations(id)` 추가.
- 있으면(NOT NULL): 이 판매는 출장/매장상담 후속 거래 → ReviewManagementCard가 **mirror 모드**로 전환
- 없으면(NULL): 일반 판매 (워크인, 수동 입력 등) → 기존 카드 동작 그대로
- 부분 인덱스 `idx_offline_sales_source_consultation` (NOT NULL일 때만)

### 1. 상담 → 판매 자동 link (신규 거래)
```
출장/매장상담 상세 (status: confirmed | in_progress | completed)
  → "판매로 처리" CTA (파란 카드, 톡상담 제외)
  → /sales/new?from_consultation={id}
  → 마운트 시 GET /api/consultation/{id} → 고객 정보 prefill
       (customerName, customerPhone, address, customerType)
  → 상단 안내 배너: "출장/매장상담에서 가져옴 · CS-..."
  → 사장님 결제·SKU만 입력 → 저장
  → POST /api/sales body.source_consultation_id 포함
  → INSERT offline_sales(source_consultation_id) → 자동 link
```

### 2. 수동 link (기존 데이터, 사장님 사후 연결)
```
sale 상세 → ReviewManagementCard 아래 작은 텍스트 "🔗 이 판매를 출장/매장상담과 연결"
  → LinkConsultationModal 열기
  → 같은 phone의 출장/매장상담 목록 (톡상담·취소 제외, 최대 10건)
  → 선택 → PATCH /api/sales/{id} { action: 'link_consultation', source_consultation_id }
  → 서버: phone_normalized 일치 검증 (오링크 방지)
  → 성공 → mirror 모드 자동 전환

해제: ReviewManagementCard 아래 "상담 연결 해제" → PATCH action='unlink_consultation'
```

### 3. ReviewManagementCard mirror 모드
| 조건 | 렌더링 |
|------|------|
| `submittedAt` 있음 | "작성 완료" 정적 라벨 (mirror 무시) |
| `linkedConsultation` 있음 (sale only) | **mirror 박스**: "원본 상담 [CS-...]에서 관리 중" + 상담 링크 |
| 그 외 | 기존 약속 토글 + 후기 요청 버튼 |

→ sale의 약속/발송 토글이 사라짐 → 원본 상담 한 곳에서만 관리 → 중복 알림톡 위험 0

### 4. 약속 대기 탭 (`/reviews`) 자동 정리
`/api/reviews/promised`의 offline_sales 쿼리에 `.is('source_consultation_id', null)` 추가.
- link 있는 sale은 약속 대기 탭에서 자동 제외 → 사장님이 같은 거래에 대해 두 행을 보지 않음
- consultation row 1개로 통합 노출

### 5. consultation 상세 — 역방향 linkedSales 노출
`useConsultation` 훅이 `offline_sales WHERE source_consultation_id = id`로 역방향 fetch (최대 5건, 취소 제외).
- 패널: "이 상담으로 판매 처리:" + 녹색 `[OS-...]` 칩 (클릭 → 판매 상세 이동)
- 풀 페이지: 녹색 강조 카드

### 회귀 안전
- link 없는 sale은 동작 변화 0 (모든 신규 prop optional + null 기본값)
- 기존 ReviewManagementCard 사용처 3곳(consultation/repair/sale) 회귀 X — sale만 linkedConsultation 전달
- DB 컬럼 추가만, 기존 데이터 그대로

상세 마이그: `supabase/migrations/070_offline_sales_source_consultation_link.sql`

---

## 리뷰 요청 분기 (2026-04-29 추가)

판매는 이미 수동 트리거(`ReviewRequestModal`) — 067에서 약속 토글 + 작성 자동 매칭 추가:

- `offline_sales.review_promised_at` (신규): 사장님 약속 ✓ 시점
- `offline_sales.review_requested_at` (기존, semantic alias of review_request_sent_at): 알림톡 발송 시점
- `offline_sales.review_submitted_at` (신규): reviews/submit가 sale_number 매칭으로 자동 기록

상세 패널의 기존 후기 요청 버튼은 공용 "리뷰 관리" 카드(`<ReviewManagementCard source="sale" />`)로 교체. `/reviews` "약속 대기" 탭에서 3 source 통합 노출.

---

## 빠른 송장 (2026-04-27 추가)

판매와 무관한 1회성 발송(샘플·간단 AS·거래처 출고·지인 발송 등)을 위해 **저장된 고객 정보로 ALPS 송장만 즉시 발급**하는 화면. 사이드바 '판매' 그룹: 판매 입력 → 판매 조회 → **빠른 송장** → B2B거래 → 계약서.

**핵심 원칙**: 신규 테이블 `manual_invoices`로 분리 → 매출 KPI / 월목표 / B2B 통계 RPC 모두 영향 없음.

```
TMS /manual-invoices 진입
  → CustomerAutocomplete로 기존 고객 검색·선택
  → 주소·연락처 미보유 고객 차단 + 고객 정보 페이지 링크
  → 품목명 직접 입력(50자, ALPS 송장 인쇄 글씨)
  → 배송메시지(선택)
  → "ALPS 송장 발급" → getNextInvoice() + bookShipment()
  → 송장번호 큰 글씨 + 복사 버튼 → ALPS 별도 PC에서 조회·출력
  → 페이지 하단 "오늘 발급한 별도 송장 N건" 미니 리스트 + 인라인 취소
```

**재사용 자산** (수정 없이 호출): `bookShipment()` / `getNextInvoice()` / `cancelShipment()` — `lib/lotte/alps-client.ts` · `<CustomerAutocomplete>` — `components/shared/customer-autocomplete.tsx`

**고객 상세 통합**: 거래 타임라인에 '빠른송장' 배지 항목 추가(취소된 건은 strikethrough + '취소됨'). 고객별 발급 이력은 customer 상세에서 조회.

**알림톡 발송**: v1 제외. 케이스 다양성(샘플/지인/AS/거래처) + 빈도 낮음 + 카카오 검수 오버헤드 감안. 향후 발급 폼에 체크박스 옵션으로 추가 가능 — DB에 `customer_id` `invoice_number` 다 쌓여 있어 데이터 손실 없음.

**파일**: API `/api/manual-invoices` (POST/GET) + `/api/manual-invoices/[id]` (DELETE) / Hook `use-manual-invoices.ts` / 폼 `quick-invoice-form.tsx` / 페이지 `(dashboard)/manual-invoices/page.tsx` / 마이그레이션 `supabase/migrations/066_manual_invoices.sql`

---

### 04-02 추가 기능
- 사이드바: '판매 입력' + '판매 조회' 분리
- 판매 조회: 주간/월간/미수금 통계 카드
- 판매 입력: 제품 목록/검색형 + 판매일 선택 + 수량 직접입력
- 판매 수정: 금액/할인/결제 편집 + 제품 추가/삭제 (내부 취소→재생성)
- 거래명세서: A4 모달 (인쇄 + 이미지 저장)
- 계약서→판매 연결 (contract_id)
- 시리얼 Race Condition 방지 (낙관적 잠금)
- raw_stock B2B 취소 시 복원
- outstanding_balance 할인 반영

---

## 1. 비즈니스 프로세스 흐름

### 오프라인 판매 등록
```
(매장 방문 고객)
  → TMS /sales/new 판매 입력
    → 판매 채널 선택 (오프라인/온라인/톡상담)
    → 고객 검색/선택 (또는 "고객 추가"로 신규 등록)
    → 카테고리별 제품 선택 (BL/TH/LO/SL)
    → 수량/할인/결제방법 입력
    → "판매 등록" → sale_number 자동 채번 (OS-YYYYMMDD-NNN)
  → TMS 재고 자동 차감 (상품별 병렬 처리)
  → 아임웹 재고 자동 차감 (재고 관리 상품만, 병렬)
  → 미결제 시 고객 미수금 자동 반영
```

### 판매 취소
```
(판매 목록 → 모달 또는 상세 페이지)
  → "판매 취소" 클릭 → 취소 사유 입력
  → PATCH /api/sales/[id] action=cancel
    1. 시리얼 복원 (sold → in_stock)
    2. 재고 복원 (products.stock_quantity += qty, 병렬)
    3. 아임웹 재고 동기화 (복원된 수량, 병렬)
    4. 미수금 차감 (미결제/부분결제였던 경우)
    5. cancelled_at / cancelled_reason / cancelled_by 기록
  → 목록에서 취소 건 dim 처리 (opacity-50 + 취소선)
```

### 결제상태 변경
```
(판매 목록 → 모달 또는 상세 페이지)
  → "결제완료로 변경" 클릭
  → PATCH /api/sales/[id] action=update_payment
    → 낙관적 업데이트 (즉시 UI 반영, 서버 실패 시 롤백)
    → 미수금 자동 조정
```

### 고객 유형별 가격 결정
```
고객 선택 시 customer_type 확인
  → dealer:  product.price_dealer > 0 → 딜러가 적용
  → academy: product.price_academy > 0 → 아카데미가 적용
  → 그 외:   product.price → 소매가 적용
장바구니 전체 재계산 (recalcCartPrices)
```

### 결제 유형 / 상태 / 채널
```
결제방식: card(카드) / cash(현금) / transfer(계좌이체) / mixed(복합)
결제상태: paid(결제완료) / unpaid(미결제) / partial(부분결제)
판매채널: offline(오프라인) / online(온라인-아임웹) / talk(톡상담)
```

---

## 2. 시스템 연동 흐름

```
[관리자]
  │
  ▼
[TMS /sales/new] ──POST──→ [/api/sales]
                               │
                               ├──→ [Supabase] offline_sales + offline_sale_items
                               ├──→ [TMS 재고] products.stock_quantity -N (병렬)
                               ├──→ [아임웹 재고] PATCH /v2/shop/products/{no} (병렬)
                               └──→ [미수금] customers.outstanding_balance 업데이트

[TMS /sales 목록] ──클릭──→ [SaleDetailModal 모달]
                               │
                               ├──→ "결제완료로 변경" → PATCH action=update_payment (낙관적)
                               └──→ "판매 취소" → PATCH action=cancel (서버 확인 필수)
                                     └──→ 시리얼 복원 + 재고 복원 + 아임웹 동기화 + 미수금 차감
```

---

## 3. 구현 완료 ✅

| 항목 | 구현 내용 |
|------|-----------|
| API | GET/POST /api/sales + PATCH /api/sales/[id] (목록+등록+취소+상태변경+메모) |
| 페이지 | /sales (목록+모달), /sales/new (입력), /sales/[id] (상세) |
| 판매 상세 모달 | 목록에서 클릭 시 모달로 즉시 조회 + 인라인 액션 (03-22) |
| 판매 취소 | 시리얼/재고/아임웹/미수금 역전 + 취소 사유 기록 (03-22) |
| 결제상태 변경 | 낙관적 업데이트 — 즉시 UI 반영, 서버 후처리 (03-22) |
| 판매 채널 | 오프라인/온라인/톡상담 칩 표시 + 입력폼 선택 (03-22) |
| 고객 | 자동완성 검색 + 신규등록 모달 (매입처 제외) |
| 가격 결정 | 고객 유형별 자동 적용: retail/online→소매가, dealer→딜러가, academy→아카데미가 |
| 재고 연동 | TMS 차감 + 아임웹 자동 동기화 (병렬 처리) |
| 회계 | COGS/마진 분석, 매출 엑셀 내보내기 |
| 미수금 | 미결제 판매 시 고객 outstanding_balance 자동 업데이트 |
| 매입처 드롭다운 | SupplierSelect 컴포넌트 |
| 성능 | SaleRow React.memo, 재고/아임웹 병렬 처리 (03-22) |
| 이카운트 이관 | CSV 분할 업로드 + 일자 정제 + 금액 보정 (03-24) |
| 후기 요청 알림톡 | 판매 상세→후기 요청 버튼→솔라피 알림톡→리뷰 페이지 (04-09) |
| 출고완료 알림톡 | 송장생성→출고완료 2단계 분리, 체크박스로 알림톡 선택 발송 (04-16) |
| 거래명세서 편집 | 미리보기에서 품목명 인라인 수정 + 납품명 자동 적용 (04-16) |
| 제품 일괄 수정 | 테이블 형태 편집 (가격/카테고리/발주명/제품군/순서) (04-16) |
| 제품 그룹핑 | 제품군(대분류)→카테고리(중분류) 2단 그룹 헤더 (04-16) |
| 제품 통합 정렬 | product_group→category→sort_order→name + 설정 연동 (04-16) |
| B2B 수금 처리 | 거래처 미수금 일괄 수금 모달 — 납품+판매 체크 선택 (04-17) |
| 납품서 품목명 편집 | B2B 납품서 미리보기 인라인 수정 + 납품명 자동 적용 (04-17) |
| 거래명세서+납품서 이미지 복사 | 클립보드 PNG 복사 버튼 (04-17) |
| 제품군 필터 칩 | 카테고리 아래 제품군(A1/A2/R1) 2차 필터 (04-17) |
| 정렬/그룹핑 정합성 | Map 기반 그룹핑 + 설정 연동 (제품군순=그룹핑, 기타=플랫) (04-17) |
| 매입처 자동 연결 | 매입품목 추가 시 products.supplier_id 자동 설정 (04-17) |
| 외화 발주 (USD/CNY) | 통화 선택 + 환율 입력 → KRW 자동 환산 (04-17) |
| 발주 부가세 변경 | 확정 후에도 포함/별도/미적용 변경 가능 (04-17) |
| 발주 수량 직접 입력 | 숫자 직접 타이핑 가능 (04-17) |
| 발주 제품 검색 | 제품명/SKU/제품군 실시간 필터 (04-17) |
| 발주 카탈로그 특징 표시 | 매입품목 features를 제품 카드에 표시 (04-17) |
| 리뷰 info API OS-* 대응 | uid가 판매번호(OS-*)일 때 offline_sales fallback 조회 (04-09) |
| 후기 webhook 변수 매핑 | consult_uid/as_uid 추가로 솔라피 템플릿 변수 정상 매핑 (04-09) |
| 판매 수정 시리얼 관리 | 수정 모달에서 시리얼 추가/삭제/변경 + 기존 시리얼 프리필 (04-11) |
| 임시제품 시리얼 | product_id 없는 임시제품에도 시리얼 입력 가능 (04-11) |
| sale_item_id 순서 보장 | insert().select()로 반환 ID 직접 사용 (created_at 순서 불보장 해결) (04-11) |
| 판매조회 취소 필터 | 전체/오늘/미수금 탭에서 취소 건 제외, 취소 탭만 표시 (04-10) |
| 고객 연락처 연동 | 판매 모달에서 customers.phone 최신 표시 (03-24) |
| 연락처 포맷 | 010-1234-5678 하이픈 자동 표시 (03-24) |
| 탭 필터 | 전체/오늘/미수금/취소 탭 바 + 탭별 건수 표시 (03-25) |
| 채널 필터 | 오프라인/온라인/톡상담 칩 필터 (03-25) |
| 기간 필터 | 전체/오늘/이번주/이번달 드롭다운 (03-25) |
| PC 테이블 뷰 | lg 이상 데이터 그리드, 모바일 카드 뷰 자동 분기 (03-25) |
| 임시 제품 입력 | 미등록 제품(빗/소모품) 이름+금액 직접 입력 (03-25) |
| 계약서 연결 | offline_sales.contract_id + 신규 계약서 알림 배너 (03-25) |
| DB 030 | offline_sales.contract_id uuid 컬럼 + 인덱스 (03-25) |

### 계약서 시스템 (03-25 리뉴얼)

| 항목 | 구현 내용 |
|------|-----------|
| 목록 탭 재편 | 전체/신규계약/전환완료/취소 + 건수 뱃지 + PC 테이블뷰 (03-25) |
| 고객 필기 캔버스 | HandwritingField — 성함/연락처/주소 S펜/터치 필기 (03-25) |
| 상담자 불러오기 | 오늘 예약 고객 모달 → 정보 자동 기입 + consultation_id 연결 (03-25) |
| 이미지 자동 캡처 | html2canvas → Supabase Storage → image_url 저장 (03-25) |
| 상세 액션 | 판매전환 / 판매전환+알림톡 버튼 + 이미지 열람 링크 (03-25) |
| DB 031 | consultation_id, handwriting 3컬럼, image_url (03-25) |

## 4. 미완료 ❌

| 항목 | 우선순위 |
|------|----------|
| 판매 모달→계약서 작성 CTA | 중간 |
| 판매전환+알림톡 활성화 (템플릿 등록 필요) | 중간 |

### 판매 수정 시 시리얼 관리 (04-11 완료)
```
(판매 조회 → 판매 건 클릭 → 수정 버튼)
  → FullEditSaleModal 열림
    → 각 아이템에 SerialPicker 표시 (기존 시리얼 프리필)
    → 시리얼 추가: 재고 선택 / 직접입력 / 자동생성
    → 시리얼 삭제: 프리필된 시리얼 해제
    → 시리얼 변경: 삭제 + 추가 조합
  → 저장 (rebuild_sale API)
    → 기존 시리얼 in_stock 복원 (원래 창고)
    → 새 시리얼 sold 전환 + sale_item_id 매핑
    → 재고 차감/복원 자동 처리
  → 준비표/시리얼조회 자동 연동
```

### 후기 요청 흐름 (04-09 완료)
```
(판매 상세 페이지)
  → "후기 요청" 버튼 클릭
  → POST /api/reviews/request (sale_id, review_type, subtype)
    → offline_sales에서 고객 정보 조회
    → 리뷰 URL 생성: page_review.html?type=TYPE&uid=OS-XXX&name=NAME&subtype=...
    → Make webhook 발송 (uid, consult_uid, as_uid, review_url 포함)
  → 솔라피 알림톡 발송 (고객에게 리뷰 링크 포함)
  → 고객이 링크 클릭 → page_review.html
    → uid(OS-*)로 /api/reviews/info 조회 (offline_sales fallback)
    → 고객 정보 자동 표시 → 후기 작성 폼
```
