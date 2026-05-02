# MAMORU TMS — 시스템 아키텍처 & 프로세스 흐름도

> 최종 수정: 2026-05-02 | **달력 관리 + 휴무 SSOT 통합 (078)** + 즉각 반영 풀 연동 (075/076) + B2B 카테고리 동적화 (074) + B2B 카탈로그 (073) + 고객 자동 매칭 (072) + 070 link

## 📌 078 (2026-05-02): 상담 달력 관리 + 휴무 SSOT 통합
- **달력 관리 화면 신설**: `/consultations/calendar` — 사장님이 4개월 달력에서 휴무 토글 (정기 요일 + 임시 날짜 통합)
- **사장님 룰 박제**: 막힘은 고객 셀프 예약에만 적용. 사장님 측 흐름(일정수동등록/시간제안)은 항상 유동
- **SSOT 통합**: 설정 → 상담 설정의 "휴무 요일" + "특별 휴무일" UI 삭제 → 달력 관리에서 통합
- **휴무 사유**: 사장님 자유 입력 (예: "결혼식"), 고객 폼에는 자동 제외 (date만 응답)
- **신규 API**: `/api/consultation/blackouts` (closed_dates), `/api/consultation/settings` (disabled_weekdays)

## 📌 075/076 (2026-04-30 심야 +3): 사장님 보고 3건 + 풀 연동 강화
- **경비 카테고리 동적화**: 설정 → 회계 → 카테고리 추가 → 즉시 /expenses 반영 (이전 hard-coded)
- **"일정 재요청" 카운트 fix**: needAction status에서 pending_admin 잘못 포함 제거 (TS fallback + RPC v3)
- **복원수리 매출 통일 (옵션 A 발생 기준)**: 모든 채널이 sale_date/created_at/delivery_date 기준 + 미입금 포함 (취소만 제외)
- **invalidate 일원화**: `lib/query/invalidate-keys.ts` — 모든 sale/repair mutation 후 대시보드 매출 즉각 갱신 (60초 지연 → 즉시)
- **staleTime**: sales-stats 60→30s, products 5분→1분
- **076 KST timezone fix**: `now() AT TIME ZONE 'Asia/Seoul'` — 매월 1일 자정~9시 매출이 전월에 잘못 합산되던 버그 fix

## 📌 운영 도구 정책 (2026-04-30 확정)
- **AppSheet**: 폐기 (2026-02-26). TMS가 전면 대체.
- **GAS (Google Apps Script)**: 폐기 (2026-04-30). Apps Script 'MAMORU_Consulting' 트리거 2개 OFF, 코드는 1주 모니터링 후 archive 폴더 이동.
- **단일 운영체계**: TMS (Vercel + Supabase + GitHub) — 모든 신규 기능은 이 스택으로만 추가
- 상세: `memory/feedback_gas_deprecated.md`

이 문서는 TMS 전체 시스템의 **마스터 레퍼런스**이다.

> **모듈별 상세**: [상담](TMS_FLOW_CONSULTATION.md) | [복원수리](TMS_FLOW_REPAIR.md) | [주문](TMS_FLOW_ORDERS.md) | [판매](TMS_FLOW_SALES.md) | [재고](TMS_FLOW_INVENTORY.md) | [회계](TMS_FLOW_ACCOUNTING.md) | [고객](TMS_FLOW_CUSTOMERS.md)
> **전체 연동 지도**: [프로세스 맵](TMS_PROCESS_MAP.md)
기능 추가/수정 시 이 문서에서 관련 파일을 찾고, 작업 후 변경사항을 반영한다.

> 모듈별 상세 흐름: [상담](TMS_FLOW_CONSULTATION.md) | [복원수리](TMS_FLOW_REPAIR.md) | [주문](TMS_FLOW_ORDERS.md) | [판매](TMS_FLOW_SALES.md) | [재고](TMS_FLOW_INVENTORY.md)

---

## §1. 시스템 개요

### 아키텍처

```
┌─────────────────────────────────────────────────────┐
│                    고객 접점                          │
├─────────────┬──────────────┬────────────────────────┤
│ GitHub Pages│ 아임웹 쇼핑몰  │ 카카오톡 알림톡 링크     │
│ (접수 폼)    │ (제품 구매)   │ (일정선택/취소/리뷰)     │
└──────┬──────┴──────┬───────┴──────────┬─────────────┘
       │             │                  │
       ▼             ▼                  ▼
┌──────────────────────────────────────────────────────┐
│              Vercel (Next.js API Routes)              │
│  /api/consultation/*  /api/repair/*  /api/lotte/*    │
│  /api/imweb/*  /api/sales/*  /api/contracts/*        │
│  /api/cron/*  /api/reviews/*  /api/products/*        │
└──────┬──────────────┬─────────────────┬──────────────┘
       │              │                 │
       ▼              ▼                 ▼
┌────────────┐ ┌─────────────┐ ┌───────────────────┐
│  Supabase  │ │ Make.com    │ │  외부 API          │
│ (DB+Auth)  │ │ → 솔라피    │ │ 롯데택배 ALPS      │
│            │ │ → 알림톡    │ │ 아임웹 API         │
│            │ │             │ │ 카카오 Geocoder    │
└────────────┘ └─────────────┘ └───────────────────┘
```

### 기술 스택

| 계층 | 기술 | 용도 |
|------|------|------|
| Frontend (TMS) | Next.js 15 + React 19 + TailwindCSS | 관리자 대시보드 |
| Frontend (고객) | 정적 HTML (GitHub Pages) | 접수 폼, 일정 선택, 리뷰 |
| Backend | Vercel Serverless Functions | API 처리 |
| DB | Supabase (PostgreSQL) | 데이터 저장 + 인증 |
| 알림 | Make.com → 솔라피 → 카카오 알림톡 | 고객 알림 |
| 이메일 | Nodemailer (Gmail) | 관리자 알림 |
| 택배 | 롯데택배 ALPS API | 송장 생성/취소/추적 |
| 쇼핑몰 | 아임웹 API v2 + OpenAPI | 주문 동기화 |
| 지도 | 카카오 Maps SDK + REST API | 출장 주소 좌표 |

---

## §2. 상담 모듈 (Consultation)

### 상태 흐름
```
[접수] pending_admin
  ├── 매장방문 (날짜+시간 있음) → confirmed (즉시 확정)
  ├── 출장요청 → pending_admin → suggested → confirmed
  └── 톡상담 → pending_admin → in_progress → completed

공통: 어느 상태든 → cancelled / on_hold 가능
확정 후: confirmed → in_progress → completed → (리뷰 요청 알림톡)
출장 재요청: suggested → reschedule_requested → suggested (재제안)
```

### 알림톡 트리거

| 시점 | 템플릿 | Make 이벤트 |
|------|--------|------------|
| 매장방문 접수 | `confirmed` | CONFIRMED |
| 출장 접수 | `request` | CONSULT_REQUEST |
| 톡상담 접수 | `talk_received` | TALK_RECEIVED |
| 시간 제안 | `suggest` | SUGGESTED_TIMES |
| 출장 확정 (고객 선택) | `field_confirmed` | FIELD_CONFIRMED |
| 출장 지연 | `field_delayed` | FIELD_DELAYED |
| 출장 24h 리마인더 | `field_remind_24h` | FIELD_REMIND_24H |
| 출장 2h 리마인더 | `field_remind_2h` | FIELD_REMIND_2H |
| 매장 24h 리마인더 (071, 2026-04-30) | `remind24` | REMIND_24H |
| 매장 2h 리마인더 (071, 2026-04-30) | `remind2` | REMIND_2H |
| 취소 | `cancelled` / `field_cancelled` | CANCELLED / FIELD_CANCELLED |
| 톡상담 시작 | `talk_ready` | TALK_READY |
| ~~상담 완료 → 후기 요청~~ | ~~`review_request`~~ | **2026-04-30 제거** — 후기는 sale source 단일 진입점 |

### 파일 매핑

| 기능 | 파일 경로 |
|------|----------|
| **고객 접수 폼** | `projects/consulting/page_form.html` |
| **시간 선택 페이지** | `projects/consulting/page_suggest.html` |
| **일정 변경/취소** | `projects/consulting/page_change_request.html` |
| **접수 API** | `app/src/app/api/consultation/public/submit/route.ts` |
| **슬롯 조회 API** | `app/src/app/api/consultation/public/slots/route.ts` |
| **확인/취소/재요청 API** | `app/src/app/api/consultation/public/confirm,cancel,resched/route.ts` |
| **시간 제안 API** | `app/src/app/api/consultation/suggest/route.ts` |
| **상태 변경 API** | `app/src/app/api/consultation/[id]/route.ts` (PATCH) |
| **삭제 API** | `app/src/app/api/consultation/[id]/route.ts` (DELETE) |
| **수동 알림 API** | `app/src/app/api/consultation/notify/route.ts` |
| **리마인더 Cron** | `app/src/app/api/cron/send-reminders/route.ts` |
| **상태 전이 규칙** | `app/src/lib/consultation/transitions.ts` |
| **React 훅** | `app/src/hooks/use-consultations.ts` |
| **목록 컴포넌트 (매장)** | `app/src/components/consultations/store-visit-list.tsx` |
| **목록 컴포넌트 (출장)** | `app/src/components/consultations/field-request-list.tsx` |
| **목록 컴포넌트 (톡)** | `app/src/components/consultations/talk-consult-list.tsx` |
| **상세 패널** | `app/src/components/consultations/consultation-detail-panel.tsx` |
| **지도** | `app/src/components/consultations/field-request-map.tsx` |
| **시간 제안 모달** | `app/src/components/consultations/suggest-time-modal.tsx` |

### 상태별 액션 가능 범위

| 상태 | 취소 | 삭제 | 비고 |
|------|------|------|------|
| pending_admin | ✅ | ✅ | 신규접수 |
| suggested | ✅ | ✅ | 시간 제안 중 |
| confirmed | ✅ | ✅ | 확정 |
| in_progress | ✅ | ✅ | 진행 중 |
| completed | ❌ | ✅ | 완료 후 삭제만 |
| cancelled | ❌ | ✅ | 취소 후 삭제만 |
| on_hold | ✅ | ✅ | 보류 |
| reschedule_requested | ✅ | ✅ | 재요청 |

> **취소**: 고객에게 알림톡 발송 + 상태 변경. **삭제**: 알림 없이 DB 완전 제거.

### 슬롯 차단 로직 (`slots/route.ts`)
- **매장방문/톡상담**: 상담 시간(durMin, 기본 60분)만 차단
- **출장 confirmed/assigned**: `예약시간 - 60분(이동)` ~ `예약시간 + 60분(상담) + 60분(복귀)` 차단
- **출장 suggested**: suggestions 배열의 각 제안 시간에 동일 버퍼 적용
- DB 설정: `consultation_settings` 테이블 (`field_buffer_before`, `field_buffer_after`)

---

## §3. 복원수리 모듈 (Repair)

### 상태 흐름
```
[접수] intake
  → pickup_scheduled (수거 예약)
  → picked_up (수거 완료)
  → inspecting (검수 중)
  → cost_notified (비용 안내)
  → payment_confirmed (입금 확인) ← paid_at 타임스탬프 독립
  → repairing (수리 중)
  → ready_to_ship (출고 대기)
  → shipped (출고 완료) ← 송장 생성 후
  → delivered (배송 완료) ← Cron 자동 추적
  → completed (완료)

어느 상태든 → cancelled 가능
```

### 알림톡 트리거

| 시점 | 템플릿 | Make URL |
|------|--------|---------|
| 접수 | `as_received` | MAKE_WEBHOOK_URL |
| 비용 안내 (수동) | `as_cost_notice` | MAKE_REPAIR_WEBHOOK_URL |
| 입금 확인 | `as_payment_confirmed` | MAKE_REPAIR_WEBHOOK_URL |
| 출고 | `as_shipped` | MAKE_REPAIR_WEBHOOK_URL |
| 취소 | `as_cancelled` | MAKE_REPAIR_WEBHOOK_URL |
| 배송 완료 → 리뷰 | `as_review_request` | MAKE_REPAIR_WEBHOOK_URL |

### 파일 매핑

| 기능 | 파일 경로 |
|------|----------|
| **고객 접수 폼** | `projects/as/page_form.html` |
| **수리 리포트 조회** | `projects/as/page_as_report.html` |
| **접수 API** | `app/src/app/api/repair/public/submit/route.ts` |
| **상태 변경 API** | `app/src/app/api/repair/[id]/route.ts` (PATCH) |
| **삭제 API** | `app/src/app/api/repair/[id]/route.ts` (DELETE) |
| **검수 저장 API** | `app/src/app/api/repair/[id]/inspect/route.ts` |
| **송장 생성 API** | `app/src/app/api/repair/[id]/ship/route.ts` (POST) |
| **송장 취소 API** | `app/src/app/api/repair/[id]/ship/route.ts` (DELETE) |
| **수동 알림 API** | `app/src/app/api/repair/[id]/notify/route.ts` |
| **ALPS 클라이언트** | `app/src/lib/lotte/alps-client.ts` |
| **상태 전이 규칙** | `app/src/lib/repair/transitions.ts` |
| **비용 계산** | `app/src/lib/repair/cost-calculator.ts` |
| **React 훅** | `app/src/hooks/use-repairs.ts` |
| **액션 카드 (UI)** | `app/src/components/repairs/sidebar-action-card.tsx` |

### 상태별 액션 가능 범위

| 상태 | 다음 전이 | 취소 | 삭제 | 비고 |
|------|----------|------|------|------|
| intake (신규접수) | pickup_scheduled / cost_notified | ✅ | ✅ | 방문수거→수거접수, 직접발송→바로 입고 |
| pickup_scheduled (입고대기) | cost_notified | ✅ | ✅ | |
| cost_notified (비용안내) | repairing | ✅ | ✅ | paid_at 독립 관리 |
| repairing (작업중) | ready_to_ship | ✅ | ✅ | |
| ready_to_ship (출고대기) | shipped | ❌ | ✅ | 송장 생성 완료 상태, 삭제 시 ALPS 취소 시도 |
| shipped (출고완료) | delivered | ❌ | ✅ | |
| delivered (배송완료) | completed | ❌ | ✅ | |
| completed (완료) | — | ❌ | ✅ | |
| cancelled (취소) | — | ❌ | ✅ | |

> **취소 vs 삭제**: 취소는 고객에게 알림톡 발송 + 상태 변경. 삭제는 알림 없이 DB에서 완전 제거.
> **파일**: `lib/repair/transitions.ts`

### ALPS 송장 생성 흐름 (`ship/route.ts → alps-client.ts`)
1. `getNextInvoice()` — Supabase `lotte_waybill_config`에서 원자적 채번 (11자리 + mod7 체크디짓)
2. `bookShipment()` — ALPS API POST (`boxTypCd: 'A'`, `ordSct: '3'`, `fareSctCd: '03'`)
3. DB 업데이트: `invoice_number`, `status → ready_to_ship`

---

## §4. 주문관리 모듈 (Orders)

### 흐름
```
아임웹 주문 발생
  → Vercel Cron (매일) /api/cron/sync-orders → Supabase orders 테이블 upsert
  → TMS에서 송장 생성 → /api/lotte/book → ALPS 발급
  → /api/imweb/push-invoice → 아임웹에 송장 입력
  → Vercel Cron (매일) /api/cron/track-delivery → 배송 완료 자동 전환
  → 배송 완료 시 리뷰 요청 알림톡 자동 발송 (purchase_review_request)
```

### 파일 매핑

| 기능 | 파일 경로 |
|------|----------|
| **주문 동기화 Cron** | `app/src/app/api/cron/sync-orders/route.ts` |
| **배송 추적 Cron** | `app/src/app/api/cron/track-delivery/route.ts` |
| **송장 생성** | `app/src/app/api/lotte/book/route.ts` |
| **송장 취소** | `app/src/app/api/lotte/cancel/route.ts` |
| **아임웹 송장 입력** | `app/src/app/api/imweb/push-invoice/route.ts` |
| **아임웹 클라이언트** | `app/src/lib/imweb/client.ts` |
| **동기화 로직** | `app/src/lib/imweb/sync.ts` |
| **ALPS 클라이언트** | `app/src/lib/lotte/alps-client.ts` |
| **레거시 ALPS** | `app/src/lib/lotte/client.ts` |

---

## §5. 판매/계약서 모듈

### 오프라인 판매 흐름
```
TMS에서 판매 입력 (/sales/new)
  → 고객 선택 + 제품 + 시리얼 선택 (ready zone만)
  → 가격 자동 적용 (딜러→딜러가, 아카데미→아카데미가)
  → VAT 자동 계산 (카드 결제 시)
  → 시리얼 → sold 상태 전환
  → 아임웹 재고 차감 (API 연동)
```

### 출장/매장상담 → 판매 link (070, 2026-04-30 추가)
```
출장/매장상담 상세 → "판매로 처리" CTA
  → /sales/new?from_consultation={id} → 고객 정보 prefill
  → 저장 시 offline_sales.source_consultation_id 기록
  → sale 상세에 linkedConsultation 정보 칩 표시 (원본 상담 가시화)

기존 sale 수동 link: ReviewManagementCard 아래 "🔗 이 판매를 출장/매장상담과 연결"
  → LinkConsultationModal → 같은 phone 상담 목록 → 선택
  → PATCH /api/sales/{id} action='link_consultation'
  → phone 일치 검증 후 link
```
**효과**: 사장님 입력 부담 절감 + 양방향 link 가시화 (sale ↔ consultation).

**리뷰 연동 (2026-04-30 밤)**: 상담관리에서 후기 진입점 전면 제거. 후기 약속/발송은 sale 또는 repair source에서만. linkedConsultation은 정보 가시화 용도(mirror 박스 제거).

상세: `docs/TMS_FLOW_SALES.md` 070 섹션

### 전자 계약서 흐름
```
TMS에서 계약서 작성 (/contracts/new)
  → 상담자 불러오기 (오늘 매장방문/출장 고객)
  → 제품 선택 + 서명 2개 (셀러+바이어)
  → 이미지 저장 (html2canvas → Supabase Storage)
```

### 파일 매핑

| 기능 | 파일 경로 |
|------|----------|
| **판매 API** | `app/src/app/api/sales/route.ts` |
| **계약서 API** | `app/src/app/api/contracts/route.ts` |
| **계약서 이미지** | `app/src/app/api/contracts/[id]/image/route.ts` |
| **상담자 불러오기** | `app/src/components/contracts/today-consultation-picker.tsx` |

---

## §6. 제품/재고/시리얼

### 시리얼 Lifecycle
```
raw (매입 원본) → ready (마킹+포장 완료) → display (진열)
                                          → sold (판매 완료)
                                          → returned (반품)
```

### 파일 매핑

| 기능 | 파일 경로 |
|------|----------|
| **제품 API** | `app/src/app/api/products/route.ts` |
| **시리얼 API** | `app/src/app/api/serials/route.ts` |
| **재고 조정 API** | `app/src/app/api/inventory/adjust/route.ts` |
| **발주 API** | `app/src/app/api/purchasing/route.ts` |

---

## §7. 리뷰 시스템

### 흐름
```
상담 완료 / 복원수리 배송완료 / 주문 배송완료
  → 알림톡 (review_request / as_review_request / purchase_review_request)
  → 고객 리뷰 폼 (projects/consulting/page_review.html)
  → POST /api/reviews/submit → Supabase reviews 테이블
  → TMS 리뷰관리에서 승인/거절
  → 승인된 리뷰 → 아임웹 위젯 표시 (GET /api/reviews/public)
```

### 파일 매핑

| 기능 | 파일 경로 |
|------|----------|
| **리뷰 폼** | `projects/consulting/page_review.html` |
| **리뷰 위젯** | `projects/consulting/page_reviews.html` |
| **리뷰 제출 API** | `app/src/app/api/reviews/submit/route.ts` |
| **공개 리뷰 API** | `app/src/app/api/reviews/public/route.ts` |

---

## §8. 알림 시스템

### 구조
```
Vercel API → sendNotification() → Make.com 웹훅 → 솔라피 → 카카오 알림톡
Vercel API → sendAdminEmail() → Gmail SMTP → 관리자 이메일
```

### 웹훅 URL 분기

| URL | 대상 템플릿 |
|-----|-----------|
| `MAKE_WEBHOOK_URL` | 상담 13종 + 리뷰 2종 (confirmed, request, suggest 등) |
| `MAKE_REPAIR_WEBHOOK_URL` | 복원수리 5종 (as_cost_notice, as_shipped 등) |

### 파일 매핑

| 기능 | 파일 경로 |
|------|----------|
| **Make 웹훅 모듈** | `app/src/lib/notification/make-webhook.ts` |
| **Gmail 모듈** | `app/src/lib/notification/email.ts` |

### 페이로드 구조 (GAS 호환)
```json
{
  "_meta": { "ts": "ISO", "version": "tms-2.2", "func": "CONFIRMED", "trigger": "tms" },
  "topic": "alrimtalk",
  "template": "confirmed",
  "event": "CONFIRMED",
  "name": "고객명",
  "phone": "01012345678",
  "channel": "kakao",
  "sms_fallback": false,
  "...추가 데이터": "..."
}
```

---

## §9. Cron 작업 (vercel.json)

| 경로 | 스케줄 | 기능 |
|------|--------|------|
| `/api/cron/send-reminders` | `*/10 * * * *` (10분) | 상담 24h/2h 리마인더 알림톡 |
| `/api/cron/sync-orders` | `0 0 * * *` (매일) | 아임웹 주문 동기화 |
| `/api/cron/track-delivery` | `0 0 * * *` (매일) | 배송 완료 자동 전환 |

인증: `Authorization: Bearer ${CRON_SECRET}`

---

## §10. 환경변수

| 변수 | 모듈 | 설명 |
|------|------|------|
| `NEXT_PUBLIC_SUPABASE_URL` | Core | Supabase 프로젝트 URL |
| `NEXT_PUBLIC_SUPABASE_ANON_KEY` | Core | Supabase 공개 키 |
| `SUPABASE_SERVICE_ROLE_KEY` | Core | Supabase 서비스 키 (서버 전용) |
| `MAKE_WEBHOOK_URL` | 알림 | 상담 알림톡 Make 웹훅 |
| `MAKE_REPAIR_WEBHOOK_URL` | 알림 | 복원수리 알림톡 Make 웹훅 |
| `GMAIL_USER` | 알림 | Gmail 발송 주소 |
| `GMAIL_APP_PASSWORD` | 알림 | Gmail 앱 비밀번호 |
| `ADMIN_EMAIL` | 알림 | 관리자 수신 이메일 |
| `LOTTE_API_URL` | 택배 | ALPS 송장 생성 URL |
| `LOTTE_CANCEL_API_URL` | 택배 | ALPS 취소 URL |
| `LOTTE_TRACK_API_URL` | 택배 | ALPS 추적 URL |
| `LOTTE_CLIENT_KEY` | 택배 | ALPS 인증 JWT |
| `LOTTE_JOBCUSTCD` | 택배 | 고객 코드 |
| `IMWEB_API_KEY` | 쇼핑몰 | 아임웹 API 키 |
| `IMWEB_API_SECRET` | 쇼핑몰 | 아임웹 시크릿 |
| `NEXT_PUBLIC_KAKAO_MAP_KEY` | 지도 | 카카오 JS 앱키 |
| `KAKAO_REST_API_KEY` | 지도 | 카카오 REST API 키 (Geocoding) |
| `CRON_SECRET` | Cron | 스케줄 작업 인증 토큰 |

---

## §11. DB 테이블 구조 (핵심)

| 테이블 | 용도 | 주요 컬럼 |
|--------|------|----------|
| `consultations` | 상담 접수 | status, consultation_type, visit_date, visit_time, unique_id, suggestions |
| `consultation_history` | 상담 이력 | consultation_id, from_status, to_status |
| `repairs` | 복원수리 | status, as_id, paid_at, invoice_number, shipped_at |
| `repair_history` | 복원수리 이력 | repair_id, from_status, to_status |
| `repair_inspections` | 검수 데이터 | repair_id, blade_condition, photos |
| `orders` | 온라인 주문 | imweb_order_no, status, invoice_number |
| `order_items` | 주문 항목 | order_id, product_id, qty |
| `customers` | 고객 | name, phone, customer_type, source |
| `products` | 제품 | sku, price_retail, price_dealer, price_academy |
| `product_serials` | 시리얼 | serial_number, status, warehouse_zone |
| `offline_sales` | 오프라인 판매 | customer_id, total, channel |
| `contracts` | 전자 계약서 | customer_id, status, signature_data |
| `reviews` | 리뷰 | source_type, rating, status |
| `consultation_settings` | 상담 설정 | start_hour, end_hour, field_buffer_before/after |
| `lotte_waybill_config` | 송장번호 관리 | current_number, end_number |
| `closed_dates` | 휴무일 | date |

마이그레이션 파일: `app/supabase/migrations/001_~042_*.sql` (42개)

---

## §12. 고객 페이지 (GitHub Pages)

**Base URL**: `https://bsm-pixel.github.io/mamoru/`

| 파일 | 용도 | 호출 API |
|------|------|---------|
| `consulting/page_form.html` | 상담 접수 폼 | `/api/consultation/public/submit, slots, settings` |
| `consulting/page_suggest.html` | 출장 시간 선택 | `/api/consultation/public/suggest-data, confirm` |
| `consulting/page_change_request.html` | 일정 변경/취소 | `/api/consultation/public/cancel, resched, reservation` |
| `consulting/page_review.html` | 리뷰 작성 폼 | `/api/reviews/submit` |
| `consulting/page_reviews.html` | 리뷰 위젯 (아임웹) | `/api/reviews/public` |
| `consulting/page_result.html` | 알림톡 결과 메시지 | 없음 (정적) |
| `consulting/page_guide.html` | 서비스 안내 | 없음 (정적) |
| `consulting/page_diag.html` | 가위 진단 (13문항) | 없음 (정적) |
| `as/page_form.html` | 복원수리 접수 폼 | `/api/repair/public/submit, holidays` |
| `as/page_as_report.html` | 수리 현황 조회 | `/api/repair/report` |
| `as/page_guide.html` | 복원수리 안내 | 없음 (정적) |
| `main/page_main.html` | 메인 홈페이지 | `/api/reviews/public` |

**Vercel 도메인**: `https://app-eta-sandy-75.vercel.app`

---

## §13. 현황 체크리스트

### 완료 ✅ (04-01)
- [x] 상담 모듈 GAS 완전 제거 (접수/알림톡/Gmail 모두 Vercel)
- [x] 복원수리 접수 Vercel 직접 처리 + Gmail 관리자 알림
- [x] 알림톡 Make 웹훅 정상 작동 (Make 시나리오 ON 확인)
- [x] ALPS 송장 생성/취소 정상 (boxTypCd 수정)
- [x] 출장 슬롯 차단 (confirmed + suggested 버퍼 60/60)
- [x] 카카오 Geocoder 좌표 자동 세팅 (프론트+서버)
- [x] 상담/복원수리 건 삭제 기능 (알림 없이)
- [x] 리뷰 시스템 (접수→승인→위젯)
- [x] 리마인더 중복 발송 방지 + KST 타임존
- [x] 복원수리 주소 필드명 정합 (address/address_detail)
- [x] 에러 메시지 [object Object] 해결 (typeof 체크 6곳)
- [x] 삭제 시 사진 파일 정리 (repair_photos + Storage)
- [x] 달력 suggested 표시
- [x] .env.local 미사용 변수 정리

### 04-01 판매관리 수정 ✅
- [x] 시리얼 Race Condition 낙관적 잠금 (eq status=in_stock)
- [x] previous_zone DB 마이그레이션 (043)
- [x] contract_id 계약서→판매 연결 (API+폼+훅)
- [x] raw_stock B2B 취소 시 복원
- [x] outstanding_balance 할인 반영
- [x] 에러 핸들링 use-sales.ts (instanceof Error)
- [x] 시리얼 수량 서버 검증
- [x] 누락 DB 컬럼 일괄 추가 (payment_detail, supply/vat, sale_channel, customer_type, contract_id)

### 미완료 / 확인 필요 ❌
- [ ] Gmail 환경변수 (GMAIL_USER, GMAIL_APP_PASSWORD) Vercel 설정
- [ ] 계약서 알림톡 템플릿 솔라피 등록 (`contracts/notify` 주석 상태)
- [ ] GAS 배포 보관처리 (2주 병렬 운영 후)
- [ ] 실전 시뮬레이션 테스트 A~F (TODO.md 참조)
- [ ] 판매관리 IA 개편 (TODO.md 상세 플랜 참조)
- [ ] 네이버 리뷰 160개 CSV 일괄 등록
- [ ] 아임웹 리뷰 위젯 코드위젯 삽입

---

## 부록: API 라우트 전체 목록

<details>
<summary>클릭하여 펼치기 (70+ routes)</summary>

### 상담 (15)
- `POST /api/consultation/public/submit` — 접수 (공개, 알림톡+Gmail)
- `GET /api/consultation/public/slots` — 가용 슬롯 (공개)
- `GET /api/consultation/public/settings` — 설정 (공개)
- `GET /api/consultation/public/suggest-data` — 제안 데이터 (공개)
- `GET /api/consultation/public/reservation` — 예약 조회 (공개)
- `GET /api/consultation/public/confirm` — 고객 확정 (공개, 알림톡)
- `GET /api/consultation/public/cancel` — 고객 취소 (공개, 알림톡)
- `GET /api/consultation/public/resched` — 고객 재요청 (공개, Gmail)
- `GET,PATCH,DELETE /api/consultation/[id]` — 조회/수정/삭제 (인증)
- `POST /api/consultation/suggest` — 시간 제안 (인증, 알림톡)
- `POST /api/consultation/notify` — 수동 알림 (인증, 알림톡)
- `POST /api/consultation/delay` — 지연 안내 (인증, 알림톡)
- `POST /api/consultation/assign` — 딜러 배정 (인증)
- `POST /api/consultation/backfill-coords` — 좌표 일괄 채우기 (인증)
- `POST /api/consultation/sync` — GAS 동기화 (듀얼 인증)

### 복원수리 (10)
- `POST /api/repair/public/submit` — 접수 (공개, 알림톡)
- `GET /api/repair/public/holidays` — 휴무일 (공개)
- `GET,PATCH,DELETE /api/repair/[id]` — 조회/수정/삭제 (인증)
- `POST /api/repair/[id]/inspect` — 검수 저장 (인증)
- `POST,DELETE /api/repair/[id]/ship` — 송장 생성/취소 (인증)
- `POST /api/repair/[id]/notify` — 수동 알림 (인증, 알림톡)
- `GET /api/repair/report` — 수리 리포트 (인증)
- `POST /api/repair/photos` — 사진 업로드 (인증)
- `POST /api/repair/sync` — 동기화 (인증)

### Cron (3)
- `GET /api/cron/send-reminders` — 리마인더 (Cron, 알림톡)
- `GET /api/cron/sync-orders` — 주문 동기화 (Cron)
- `GET /api/cron/track-delivery` — 배송 추적 (Cron)

### 주문/아임웹 (4)
- `POST /api/imweb/sync` — 주문 동기화 (인증)
- `POST /api/imweb/push-invoice` — 송장 입력 (인증)
- `POST /api/imweb/sync-products` — 제품 동기화 (인증)
- `POST /api/imweb/orders/[id]/review-request` — 리뷰 요청 (인증, 알림톡)

### 롯데택배 (4)
- `POST /api/lotte/book` — 송장 생성 (인증)
- `POST /api/lotte/cancel` — 송장 취소 (인증)
- `GET /api/lotte/check-cancel` — 취소 확인 (인증)
- `GET /api/lotte/track` — 배송 추적 (인증)

### 판매/계약서/고객/제품/시리얼/재고/발주/리뷰/딜러/설정
- 각 모듈별 CRUD 라우트 (상세는 해당 모듈 섹션 참조)

</details>