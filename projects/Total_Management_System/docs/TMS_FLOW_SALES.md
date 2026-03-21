# 판매관리 프로세스 흐름도
> 최종 업데이트: 2026-03-21

---

## 1. 비즈니스 프로세스 흐름

### 오프라인 판매 등록
```
(매장 방문 고객)
  → TMS /sales/new 판매 입력
    → 고객 검색/선택 (또는 "고객 추가"로 신규 등록)
    → 카테고리별 제품 선택 (BL/TH/LO/SL)
    → 수량/할인/결제방법 입력
    → "판매 등록" → sale_number 자동 채번 (OS-YYYYMMDD-NNN)
  → TMS 재고 자동 차감
  → 아임웹 재고 자동 차감 (재고 관리 상품만)
  → 미결제 시 고객 미수금 자동 반영
```

### 결제 유형 / 상태
```
결제방식: card(카드) / cash(현금) / transfer(계좌이체) / mixed(복합)
결제상태: paid(결제완료) / unpaid(미결제) / partial(부분결제)
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
                               ├──→ [TMS 재고] products.stock_quantity -N
                               ├──→ [아임웹 재고] PATCH /v2/shop/products/{no} (자동)
                               └──→ [미수금] customers.outstanding_balance 업데이트
```

---

## 3. 구현 완료 ✅

| 항목 | 구현 내용 |
|------|-----------|
| API | GET/POST /api/sales (목록+등록) |
| 페이지 | /sales (목록), /sales/new (입력), /sales/[id] (상세) |
| 고객 | 자동완성 검색 + 신규등록 모달 |
| 재고 연동 | TMS 차감 + 아임웹 자동 동기화 |
| 회계 | COGS/마진 분석, 매출 엑셀 내보내기 |
| 미수금 | 미결제 판매 시 고객 outstanding_balance 자동 업데이트 |
| 매입처 드롭다운 | SupplierSelect 컴포넌트 |

## 4. 미완료 ❌

| 항목 | 우선순위 |
|------|----------|
| 판매 수정/삭제 기능 | 낮음 |
