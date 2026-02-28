# 주문관리 프로세스 흐름도
> 최종 업데이트: 2026-02-28

---

## 1. 비즈니스 프로세스 흐름

### 주문 수신 → 배송 흐름
```
[아임웹] 고객 결제
  → TMS Cron 동기화 (또는 수동 동기화 버튼)
    → GET /v2/shop/orders → Supabase upsert
  → [status: pay_done] (결제완료)
  → (관리자) 아임웹에서 "배송대기" 처리 (수동 1회)
  → (관리자) TMS에서 "송장 생성" 클릭
    → ALPS 접수 → 12자리 운송장 생성
    → 아임웹 자동 송장 입력 (STANDBY 상태에서)
  → [status: shipping] (배송중)
  → 배송 완료 → [status: delivered]
```

### 취소 흐름
```
(관리자) TMS "송장 취소" 클릭
  → [status: cancel_pending] (TMS에서만 상태 변경)
  → (관리자) ALPS 사이트에서 집하취소 (수동)
  → TMS "ALPS 취소 확인" 버튼 → 상태 자동 감지
    → ALPS 취소됨 → [status: cancelled]
  → 아임웹 상태변경은 v2 API 미지원 → 수동 처리
```

### 동기화 보호 규칙
```
아임웹 Cron 동기화 시 TMS-managed 상태는 덮어쓰기 방지:
- cancel_pending (TMS에서 취소 진행 중)
- shipping (TMS에서 송장 생성 완료)
→ 이 상태들은 아임웹 데이터로 덮어쓰지 않음
```

### 상태 목록
```
pay_wait     — 입금대기
pay_done     — 결제완료
preparing    — 상품준비중
shipping     — 배송중
delivered    — 배송완료
cancel_pending — 취소진행중 (TMS 전용)
cancelled    — 취소완료
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

### 아임웹 API 제약
| 기능 | 가능여부 | 비고 |
|------|----------|------|
| 주문 조회 | ✅ | GET /v2/shop/orders |
| 상품주문 조회 | ✅ | GET /v2/shop/orders/{no}/prod-orders |
| 송장 입력 | ✅ | PATCH /v2/shop/prod-orders/{no}/invoice (STANDBY 이상) |
| 주문 상태 변경 | ❌ | code -99 |
| 주문 취소/환불 | ❌ | 수동 처리 |
| 웹훅 | ❌ | 미지원 |

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

| 항목 | 의존성 | 우선순위 |
|------|--------|----------|
| 온라인 주문 → 이카운트 매출전표 자동 연동 | 이카운트 API (검증완료), 제품→품목 코드 매핑 | 중간 |
| 배송완료 자동 전환 (ALPS 추적 → delivered) | Cron + queryStatus | 낮음 |
| Vercel Cron 5~10분 간격 자동 폴링 활성화 | vercel.json cron 설정 | 낮음 |
| 동기화 상태 대시보드 표시 (마지막 시간, 에러) | 없음 | 낮음 |

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
| `app/src/lib/imweb/client.ts` | 아임웹 v2 API 클라이언트 |
| `app/src/lib/imweb/sync.ts` | 주문 동기화 로직 (증분+보호) |
| `app/src/lib/imweb/types.ts` | 아임웹 타입 정의 |
| `app/src/lib/lotte/client.ts` | 롯데택배 ALPS 클라이언트 (book/cancel/track) |
| `app/src/lib/lotte/types.ts` | 롯데택배 타입 정의 |
| `app/src/hooks/use-orders.ts` | React Query 훅 5개 |

### DB 마이그레이션
| 파일 | 설명 |
|------|------|
| `app/supabase/migrations/001_orders.sql` | orders + order_items + sync_log + waybill_counter |
