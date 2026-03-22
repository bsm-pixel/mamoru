# 판매관리 프로세스 흐름도
> 최종 업데이트: 2026-03-22

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

## 4. 미완료 ❌

| 항목 | 우선순위 |
|------|----------|
| 판매 수정 기능 | 낮음 |
