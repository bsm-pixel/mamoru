# 판매관리 프로세스 흐름도
> 최종 업데이트: 2026-04-09 | 동적 단가 그룹 + 복원수리 카테고리(RS) + 후기 요청 + 매출 발생주의 + 필터 IA 개선
>
> **마스터 문서**: [TMS_SYSTEM_ARCHITECTURE.md](TMS_SYSTEM_ARCHITECTURE.md) §5, [TMS_PROCESS_MAP.md](TMS_PROCESS_MAP.md) 참조

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
| 리뷰 info API OS-* 대응 | uid가 판매번호(OS-*)일 때 offline_sales fallback 조회 (04-09) |
| 후기 webhook 변수 매핑 | consult_uid/as_uid 추가로 솔라피 템플릿 변수 정상 매핑 (04-09) |
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
| 판매 수정 기능 | 낮음 |
| 판매 모달→계약서 작성 CTA | 중간 |
| 판매전환+알림톡 활성화 (템플릿 등록 필요) | 중간 |

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
