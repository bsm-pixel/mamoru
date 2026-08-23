# 주문관리 프로세스 흐름도
> 최종 업데이트: 2026-08-23 (마이그 128/129 — 배송대기 3단계 + 아임웹 배송상태 자동 역동기)

---

## 1. 비즈니스 프로세스 흐름

### 주문 수신 → 배송 흐름 (2026-08-23 자동화 완성)
```
[아임웹] 고객 결제
  → TMS Cron 동기화(매일 09시) 또는 웹훅(실시간) 또는 수동 동기화 버튼
    → GET /v2/shop/orders → Supabase upsert
  → [status: pay_done] (결제완료)
  → (관리자) TMS에서 "송장 생성" 클릭 한 번  ← 아임웹 수동 개입 불필요!
    → ALPS 접수 → 운송장 생성
    → 아임웹 자동 역동기: place(배송대기 전환) → invoice(송장번호 등록)  ※ 전 품목 루프
  → [status: ready_to_ship] (배송대기) / 아임웹도 배송대기(STANDBY)
  → (기사) 집하 스캔 → track-delivery 크론(30분)이 ALPS 집하(godsStatCd 10) 감지
    → [status: shipping] (배송중) + shipped_at 기록
    → 아임웹 자동 역동기: send(발송처리) → 아임웹 배송중(DELIVERING)  ※ 전 품목 루프
  → 배송 완료 → track-delivery 크론이 ALPS 배달완료 감지 → [status: delivered]
    → 아임웹은 실제 롯데 송장 추적으로 자체 배송완료 처리 (complete 강제호출 불필요)
```

> **핵심(2026-08-23 실측 확정)**: 아임웹 v2 API로 배송상태 역동기가 **가능**하다.
> `PATCH /v2/shop/prod-orders/{품목주문번호}/{place|invoice|send}?order_version=v2`
> place=배송대기 / invoice=송장등록(배송대기 유지) / send=배송중. **전 품목주문(prod-order) 루프 필수**
> (하나만 처리 시 아임웹 "부분출고"). 과거 "code -99 미지원"은 오해 — 엔드포인트 형식 오류였음.
> 송장 발급 ≠ 출고: 복원수리·납품과 동일 리듬으로 통일(집하 전=배송대기, 집하 후=배송중).

### 취소 흐름
```
(관리자) TMS "송장 취소" 클릭
  → [status: cancel_pending] (TMS에서만 상태 변경)
  → (관리자) ALPS 사이트에서 집하취소 (수동)
  → TMS "ALPS 취소 확인" 버튼 → 상태 자동 감지
    → ALPS 취소됨 → [status: cancelled]
  → 아임웹 상태변경은 v2 API 미지원 → 수동 처리
```

### 재고 연동 (04-15 업데이트 — 아임웹 자동 동기화 완성)
```
주문 동기화 시 재고 자동 처리:
- 결제완료(pay_done) 이상 최초 진입 → TMS 재고 -N (stock_deducted=true) + 아임웹 재고 -N
- 취소/환불 → TMS 재고 +N 복구 (stock_deducted=false) + 아임웹 재고 +N
- 중복 차감 방지: orders.stock_deducted 플래그로 관리

재고 동기화 방식 (04-15):
- 새 OpenAPI: PATCH /products/{prodNo}/stock-info (증감값 delta 방식)
- 인증: OAuth2 Bearer Token (DB system_settings에서 자동 관리)
- 모든 재고 변동 시 자동 호출 (주문/판매/취소/반품/납품/발주/재고조정)
```

### 동기화 보호 규칙
```
아임웹 Cron 동기화 시 TMS-managed 상태는 덮어쓰기 방지:
- cancel_pending (TMS에서 취소 진행 중)
- ready_to_ship (TMS에서 송장 생성 완료 = 배송대기)
- shipping (TMS에서 집하 감지 = 배송중)
→ 이 상태들은 아임웹 데이터로 덮어쓰지 않음 (sync.ts TMS_MANAGED_STATUSES)
```

### 재고 차감 상태 (STOCK_DEDUCT_STATUSES — 재고 정합성)
```
pay_done · preparing · ready_to_ship · shipping · delivered · confirmed → 재고 차감 유지
cancelled · refunded → 재고 복구
※ ready_to_ship(배송대기)도 결제완료 이상이므로 반드시 차감 상태에 포함 (누락 시 재고 오복구)
```

### 상태 목록
```
pay_wait      — 입금대기
pay_done      — 결제완료
preparing     — 상품준비중
ready_to_ship — 배송대기 (128 신규: 송장생성 완료, 집하 전)
shipping      — 배송중 (집하 후)
delivered     — 배송완료
cancel_pending — 취소진행중 (TMS 전용)
cancelled     — 취소완료
```

---

## 2. 시스템 연동 흐름

```
[고객] ──결제──→ [아임웹]
                    │
          ┌─────── Cron (주기적) / 수동 동기화
          ▼
   [TMS /api/imweb/sync]
   GET /v2/shop/orders
          │
          ▼
   [Supabase DB] (orders + order_items + sync_log)
          │
          ▼
   [TMS /orders UI]
          │
     ┌────┴─────┐
     ▼          ▼
  송장 생성   송장 취소
     │          │
     ▼          ▼
[/api/lotte/book]  [/api/lotte/cancel]
     │                   │
     ▼                   ▼
[롯데택배 ALPS]     [status: cancel_pending]
(apiSndOutSingle)        │
     │              ALPS 수동 취소
     ▼                   │
[아임웹 송장 입력]        ▼
PATCH /v2/shop/     [/api/lotte/check-cancel]
prod-orders/invoice  queryStatus → cancelled
```

### 운송장 번호 생성
```
Supabase waybill_counter (싱글턴 테이블)
  → next_waybill() 함수 (atomic increment)
  → 범위: 31765377481 ~ 31765380480 (3000개)
  → 체크디짓: base11 % 7
  → 결과: 12자리 문자열
```

### 아임웹 v2 API (주문/송장용)
| 기능 | 가능여부 | 비고 |
|------|----------|------|
| 주문 조회 | ✅ | GET /v2/shop/orders |
| 상품주문 조회 | ✅ | GET /v2/shop/orders/{no}/prod-orders |
| 배송대기 전환 | ✅ | PATCH /v2/shop/prod-orders/{no}/place?order_version=v2 (결제완료→배송대기, 2026-08-23 실측) |
| 송장 입력 | ✅ | PATCH /v2/shop/prod-orders/{no}/invoice?order_version=v2 (배송대기 이상, 등록해도 배송대기 유지) |
| 발송 처리 | ✅ | PATCH /v2/shop/prod-orders/{no}/send?order_version=v2 (배송대기→배송중, 집하 시 호출) |
| 강제 배송완료 | (미사용) | PATCH /v2/shop/prod-orders/{no}/complete — 실 롯데송장이면 아임웹 자동완료라 불필요 |
| 상품 목록 조회 | ✅ | GET /v2/shop/products |
| 상품 재고 수정 | ❌ | v2로는 200 반환하지만 실제 미반영 → 새 OpenAPI 사용 |
| ~~주문 상태 변경 code -99~~ | ✅ 정정 | 과거 "-99 미지원"은 엔드포인트 형식 오류였음(위 place/send로 가능). order_version=v2 + 전 품목 루프 필수 |

### 아임웹 새 OpenAPI (재고+주문 상태용, 04-15 연동 완료)
| 기능 | 가능여부 | 비고 |
|------|----------|------|
| 상품 재고 수정 | ✅ 구현 | PATCH /products/{prodNo}/stock-info (delta 증감값) |
| 주문 취소 접수 | 🔜 예정 | PATCH /orders (취소 접수 요청) |
| 송장 등록/수정/삭제 | 🔜 예정 | POST/PATCH/DEL /orders/{no}/invoice |
| 배송 처리 | 🔜 예정 | PATCH /orders (배송 상태 변경) |
| 상품 등록/수정 | 🔜 예정 | POST/PUT /products |
| 인증 | OAuth2 | Bearer Token, DB 자동 관리, refreshToken 갱신 |
| 환경변수 | — | IMWEB_OPENAPI_KEY (clientId), IMWEB_OPENAPI_SECRET |
| 콜백 | — | /api/imweb/oauth/callback |
| OAuth URL | — | openapi.imweb.me/oauth2/authorize?...&siteCode=S20250825bc9b09c7146df |

### 롯데택배 ALPS API
| 기능 | 엔드포인트 | 비고 |
|------|------------|------|
| 접수 (송장생성) | POST `/api/pid/cus/714a/apiSndOutSingle` | invNo + checkDigit |
| 집하취소 | POST cancel endpoint | canCd='01', canDtlCd='19' |
| 배송추적 | GET `/api/pid/cus/714a/custmer-view-tracking` | invNo 기반 |

---

## 3. 구현 완료 ✅

### API Routes (7개)
| 엔드포인트 | 메서드 | 기능 |
|------------|--------|------|
| `/api/imweb/sync` | POST | 아임웹 주문 수동 동기화 |
| `/api/imweb/push-invoice` | POST | 아임웹 송장 재입력 |
| `/api/lotte/book` | POST | 송장 생성 (ALPS + 아임웹 자동 연동) |
| `/api/lotte/track` | GET | 배송 추적 |
| `/api/lotte/cancel` | POST | 송장 취소 (소프트) |
| `/api/lotte/check-cancel` | POST | ALPS 취소 확인 + 자동 상태 반영 |
| `/api/cron/sync-orders` | GET | Vercel Cron 자동 동기화 |

### 페이지 (3개)
- `/orders` — 주문 목록 (상태 탭, 검색, 결제 칩, 메모 말줄임, 페이징)
- `/orders/dashboard` — 주문 대시보드 (파이프라인 바, 통계 4개, 긴급 리스트)
- `/orders/[id]` — 주문 상세 (주문자/배송/상품/결제/배송관리)

### 컴포넌트
- `invoice-modal.tsx` — 송장 생성 모달 (수신자 정보 자동 채움)

### 동기화 로직
- 증분 동기화: 마지막 sync 시점 이후 변경분만 가져옴 (기본 30일)
- TMS-managed 상태 보호 (cancel_pending, shipping)
- sync_log 테이블에 동기화 이력 기록

### DB 테이블
- `orders` — 주문 메인 (아임웹 번호, 주문자/수신자, 결제, 송장, 상태)
- `order_items` — 주문 상품 (제품명, 옵션, 수량, 가격)
- `sync_log` — 동기화 이력
- `waybill_counter` — 운송장 번호 생성기 (싱글턴)

---

## 4. 미완료 ❌

| 항목 | 우선순위 |
|------|----------|
| 동기화 상태 대시보드 표시 (마지막 시간, 에러) | 낮음 |

---

## 5. 핵심 파일 맵

### TMS API
| 파일 | 설명 |
|------|------|
| `app/src/app/api/imweb/sync/route.ts` | 아임웹 주문 동기화 |
| `app/src/app/api/imweb/push-invoice/route.ts` | 아임웹 송장 재입력 |
| `app/src/app/api/lotte/book/route.ts` | 송장 생성 (ALPS) |
| `app/src/app/api/lotte/track/route.ts` | 배송 추적 |
| `app/src/app/api/lotte/cancel/route.ts` | 송장 취소 |
| `app/src/app/api/lotte/check-cancel/route.ts` | ALPS 취소 확인 |
| `app/src/app/api/cron/sync-orders/route.ts` | Cron 자동 동기화 |

### TMS UI
| 파일 | 설명 |
|------|------|
| `app/src/app/(dashboard)/orders/page.tsx` | 주문 목록 |
| `app/src/app/(dashboard)/orders/dashboard/page.tsx` | 주문 대시보드 |
| `app/src/app/(dashboard)/orders/[id]/page.tsx` | 주문 상세 |
| `app/src/components/orders/invoice-modal.tsx` | 송장 생성 모달 |

### TMS Lib
| 파일 | 설명 |
|------|------|
| `app/src/lib/imweb/client.ts` | 아임웹 v2 API 클라이언트 — **prepareImwebDelivery(place)·updateInvoice(전 품목)·shipImwebOrder(send)** 배송상태 역동기 |
| `app/src/lib/imweb/sync.ts` | 주문 동기화 로직 (증분+보호) — STOCK_DEDUCT/TMS_MANAGED에 ready_to_ship 포함 |
| `app/src/app/api/cron/track-delivery/route.ts` | [1-A] 주문 집하 감지(ready_to_ship→shipping) + 아임웹 send back-sync |
| `app/src/lib/imweb/types.ts` | 아임웹 타입 정의 |
| `app/src/lib/lotte/client.ts` | 롯데택배 ALPS 클라이언트 (book/cancel/track) |
| `app/src/lib/lotte/types.ts` | 롯데택배 타입 정의 |
| `app/src/hooks/use-orders.ts` | React Query 훅 5개 |

### DB 마이그레이션
| 파일 | 설명 |
|------|------|
| `app/supabase/migrations/001_orders.sql` | orders + order_items + sync_log + waybill_counter |
