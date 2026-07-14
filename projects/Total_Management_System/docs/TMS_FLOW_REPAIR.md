# 복원수리 프로세스 흐름도
> 최종 업데이트: 2026-06-09 — 제품+복원수리 혼합 장바구니 통합 (B/C채널 입력 경로)
>
> **마스터 문서**: [TMS_SYSTEM_ARCHITECTURE.md](TMS_SYSTEM_ARCHITECTURE.md) §3 참조

## 2026-06-09 — B/C채널 복원수리, 제품과 한 화면에서 함께 입력 (양자택일 모드 제거)

출장상담 등에서 제품 판매 + 복원수리를 **한 번에** 입력. 집계(A+B+C, category='RS')는 그대로 — 매출·수량만 기록(물리 파이프라인 A채널 repairs 미생성, 이중집계 0).

- **B채널 (`/sales/new`)**: "제품 판매/복원수리" 토글 폐지. 제품 목록과 **복원수리 입력 카드를 동시 노출**, 한 장바구니에 제품+RS 함께 저장. (마모루/타사 분류·배송비 정책 동일)
- **C채널 (`/deliveries` 납품 모달)**: "제품 납품/복원수리" 토글 폐지. 한 납품서에 제품 + 복원수리(선택) 동시. RS는 **VAT 제외**(`lib/deliveries/totals.ts` `computeDeliveryTotals`).
- 편집 fix: `delivery-detail-panel` 편집 시 category 보존(RS 둔갑 방지). 상세: [TMS_FLOW_SALES](TMS_FLOW_SALES.md) 2026-06-09 항목.

## 2026-05-29 — 수리내역서 v2 (디지털 핀 마킹) ⭐ 박제

검수(수리내역서 작성)를 7항목 체크리스트 → **가위 사진 위 디지털 핀 마킹**으로 전환. 디자인모니터(`/design-lab`) 좌/우 실시간 데모로 확정 후 라이브 승격.

- **입력(`components/repairs/inspection-form.tsx`)**: 가위별 탭. 사진 촬영 → 위치형 핀(무뎌짐=드래그 선 / 찍힘·부품 문제·스토퍼 문제=탭 점 ✓꼭지점 / 빗살 손상=틴닝만) + 플래그(장력조절 필요·밸런스 불균형·날각 문제 = 우측 상단 표기). 가위별 **진단 및 내역(comment)** + **핀/플래그 선택 시 멘트 자동삽입**(`components/repairs/ment-linkage.ts`, 공통/종류별 문구 = `lib/repair/comment-presets.ts`).
- **저장**: `repair_inspections.photo_marks`(jsonb)에 점/선(x2,y2)/플래그(flag:true) 한 배열로. 가위별 멘트 = `repair_inspections.comment`(마이그레이션 **097**). 7항목 컬럼(blade_*/comb/tension/parts/stopper)은 비파괴 보존(미사용).
- **고객(`projects/as/page_as_report.html`)**: 가위 사진 위 점(✓)·선·플래그 오버레이 + 가위별 "진단 및 내역"(항목별 `–`). 구 데이터 `repairs.admin_note`는 "추가 안내"로 하위호환.
- **버그 fix(동반)**: 모바일 카메라 촬영 후 dialog 임의 close로 목록 튕기던 문제 → `components/ui/modal.tsx` `preventAutoClose` 가드. 검수 화면 모바일 가로 차단 오버레이("세로로 돌려주세요").
- 공용 컴포넌트: `inspection-mark-board.tsx`(편집·controlled), `mark-overlay.tsx`(읽기전용), `inspection-marks.ts`(유형/색/타입).
- 관련 메모리: [[reference_repair_inspection_v2]]

## 2026-05-19 — 복원수리 매출 집계 기준 (회계 vs 대시보드) ⭐ 박제

복원수리 매출은 **A 접수(repairs) + B 판매RS(offline_sale_items) + C 납품RS(delivery_items)** 3채널 합산.
B·C는 두 화면 완전 동일. **A채널만 화면 목적에 따라 기준이 다름** (의도된 설계, 사장님 결정 2026-05-19):

| 구분 | 복원수리 대시보드 `/repairs` | 회계 `/reports` |
|---|---|---|
| **목적** | 이번달 들어온 일감 (운영 현황) | 실제 받은 돈 (재무) |
| **A채널 시점** | 접수 발생 (`created_at`) — 미입금 포함 | 입금 완료 (`paid_at`) |
| **A채널 금액** | `total_amount` 안분 (실제청구액) | `total_amount` (실제청구액) |

→ **금액 계산식은 통일**(실제청구액 = 수리비+배송비+가공비), **시점만 다름**. 두 화면 숫자가 약간 다른 건 정상 (미입금/시점차 = 대시보드엔 포함, 회계엔 미포함).

### A채널 마모루/타사 안분 (2026-05-19, `088` RPC + use-dashboard-stats)
- `repairs.total_amount`(건당 합산값)를 단가 비중으로 마모루/타사 분리:
  - 마모루분 = `ROUND(total_amount × (qty_mamoru×10000) / (qty_mamoru×10000 + qty_other×20000))`
  - 타사분 = `total_amount − 마모루분` (반올림 잔액 흡수 → 분리합 = total_amount 정확)
- **이전(080)은 고정단가**(자루×만원)였음 → 배송비·날변형 가공비 누락. 안분으로 실제청구액 반영.
- qty 수정 시 `calcTotalCost`(repair-detail-card)가 `total_amount` 자동 재계산 → 대시보드·회계 즉시 반영.
- 자루 수(count)는 그대로 `qty_mamoru + qty_other` (배송비 무관).
- 수정 파일: `use-dashboard-stats.ts`(/repairs), `088_hub_stats_repair_amount_prorate.sql`(/dashboard 허브 RPC). 회계 `summary/route.ts`는 원래 total_amount라 변경 없음.

## 2026-05-13 후속 — 배송비 처리
- 복원수리 입력 시 "배송비 3,000원" 항목(`/sales/new` 복원수리 모드 토글 / `/deliveries` "+B2B수리" 버튼)은 `category='RS'`, product_name='배송비' 로 저장.
- 배송비는 **복원수리 매출 금액엔 포함** (접수시스템 `repairs.shipping_fee` 와 정책 통일), **"자루 수" 카운트에선 제외** — 대시보드/회계/RPC 080 의 자루 합산 서브쿼리에 `product_name <> '배송비'`.

## 2026-05-12 — 복원수리 매출 3채널 통합 + 입력 경로 정리

복원수리 매출이 3개 채널에서 발생 → 대시보드·회계 모두 통합 집계:
- **A채널 (접수시스템 `repairs`)**: 고객이 접수페이지로 가위 보냄 → 입고·검수·수리·출고. 단가는 마모루 1만 / 타사 2만 고정(대시보드) 또는 service_cost+shipping_fee 실제 금액(회계, paid_at 기준).
- **B채널 (판매시스템 `offline_sale_items` category='RS')**: 매장에서 즉석 복원수리 판매. 입력: `/sales/new` "복원수리" 모드 (마모루 N자루 / 타사 M자루 + 단가, 기본 1만/2만). product_name 고정 "복원수리 (마모루)" / "복원수리 (타사)" → `includes('타사')` 로 마모루/타사 분류. parent 판매의 customer_type 이 dealer/academy 면 B2B 버킷으로.
- **C채널 (납품 `delivery_items` category='RS')**: B2B 거래처 복원수리. 입력: `/deliveries` "+B2B수리" 모드. 마모루/타사 구분 없음(자루당 거래처 단가). 거래처 마스터의 `default_repair_price` 가 있으면 거래처 선택 시 단가 자동 채움(없으면 기본 8천).
- 대시보드 `monthRepairAmount` = A + B + C 전체, `monthRepairCount` = 세 채널 자루(quantity) 합. `monthRepairMamoru/Other` = A + B채널 B2C분, `monthRepairB2B` = B채널 B2B분 + C채널 전체.
- 회계 리포트(2-D): `repair_sales.total` = A(repairs, paid_at, 실제금액) + B(offline RS) + C(delivery RS), 채널별 분해 표시.
- 거래처 단가: `customers.default_repair_price` (마이그레이션 079). 고객 상세 화면(딜러/아카데미)에서 입력.

## 2026-04-30 (심야) — 고객 자동 매칭/생성

복원수리 접수 시 phone 기준 자동 매칭/생성:
- `/api/repair/public/submit` (고객 폼) → `lib/customer/match-or-create.ts` 호출 → repairs.customer_id 자동 연결
- `/api/repair` POST (관리자 수기) → body.customer_id 없을 때만 자동 매칭 (사장님 customer-autocomplete 선택은 그대로)

같은 phone으로 상담/복원수리/판매 재접수 시 동일 customer_id 매칭 → 거래 이력 통합 + 단골 자동 인지.

상세: `docs/TMS_FLOW_CONSULTATION.md` "2026-04-30 심야" 섹션

---

## 리뷰 요청 분기 (2026-04-29 추가 / 2026-07-12 버그 수정)

배송완료(`delivered`) 진입 시 후기 알림톡 발송 정책. `system_settings.review.auto_request_on_completion` 토글 + **약속(`review_promised_at`) 받은 고객만** 자동 발송 (판매와 동일 정책).

🔴 **2026-07-12 (109) 버그 수정 — 자동 배송완료 건은 리뷰 알림톡이 아예 안 나가고 있었다.**
크론(`track-delivery`)이 `repairs` 를 **DB 직접 update** 해서, 발송 코드가 있는 `PATCH /api/repair/[id]` 를 우회했기 때문.
→ 사장님이 **수동으로** 상태를 바꿀 때만 발송됐고, **ALPS 자동 감지로 배송완료된 건은 발송 0건**.
→ 크론 repairs 블록에 판매와 동일한 4중 가드(토글 → 약속 → 미발송 → 전화번호)를 복제해 해결.

발송 경로 2개 (둘 다 같은 가드):
| 경로 | 트리거 |
|---|---|
| `PATCH /api/repair/[id]` | 사장님이 화면에서 상태를 delivered 로 변경 |
| `cron/track-delivery` [2] | ALPS 41/45 자동 감지 ← **109 에서 추가** |

3 timestamp: `review_promised_at` / `review_request_sent_at` / `review_submitted_at`. reviews/submit가 `repairs.as_id` 매칭으로 review_submitted_at 자동 기록. 상세 패널의 "리뷰 관리" 카드와 `/reviews` "약속 대기" 탭에서 추적.

## 출고 = 집하 자동 감지 (109, 2026-07-12)

송장 발급 시 `status='ready_to_ship'`. **롯데 기사님이 수거 스캔(ALPS `10`)하면 크론이 자동으로 `shipped` 전환 + `as_shipped` 알림톡 발송.**
`shipped_source='alps_pickup'` 으로 자동/수동 구분. 상세는 [TMS_FLOW_AUTO_DELIVERY.md](TMS_FLOW_AUTO_DELIVERY.md) 참조.

---

---

## 1. 비즈니스 프로세스 흐름

### 전체 파이프라인 (6단계)
```
신규접수 → 입고대기 → 작업중 → 출고대기 → 출고완료 → 배송완료
```

### 방문수거 흐름
```
고객 접수 (page_form.html)
  → Vercel API: /api/repair/public/submit → Supabase 저장 + as_id 채번(AS-YYYYMMDD-NNN)
  → Make 알림톡 (as_received) + TMS 동기화
  → [status: intake] (신규접수)
  → (관리자) 접수확인 → confirmed_at 기록
  → (관리자) 수거접수 완료 → [status: pickup_scheduled] (입고대기)
  → 입고 & 검수 + 비용안내 → [status: cost_notified] (작업중)
    → 알림톡 (as_cost_notice) — UI 버튼 → sendNotify API
  → (고객) 입금 → (관리자) 입금확인 → paid_at 기록
    → 알림톡 (as_payment_confirmed) — PATCH API 내부 자동
  → 수리 진행 → [status: repairing] (작업중)
  → 포장 완료 → packed_at 기록
  → 송장 생성 (ALPS) → [status: ready_to_ship] (출고대기)
  → 출고 완료 → [status: shipped] (출고완료)
    → 알림톡 (as_shipped) — PATCH API → getAutoNotifyTemplate 자동
  → 배송 완료 → [status: delivered] → [status: completed]
  → 만족도 알림톡 (as_satisfaction) — 미구현
  ※ 모든 단계에서 취소 가능 → [status: cancelled]
    → 알림톡 (as_cancelled) — PATCH API → getAutoNotifyTemplate 자동
```

### 직접발송 흐름
```
고객 접수 → [intake]
  → (관리자) 접수확인 → confirmed_at
  → 입고 & 검수 + 비용안내 → [cost_notified] (pickup_scheduled 생략)
  → 이하 방문수거와 동일
```

### 🆕 직접방문(당일수리) 흐름 — 2026-05-25 추가
> 사장님 비전: 매장 워크인. 가위 들고 와서 30분~1시간 내 당일 수리받아 감.
> 컨설팅 매장방문/출장 일정과 양방향 충돌 차단. Google Calendar 자동 박힘.
> 상세: [[project_repair_direct_visit]] 메모리

```
고객 접수 (page_form.html "직접방문" segment)
  → 방문 날짜 선택 → /api/repair/public/slots?date=&qty= → 30분 슬롯 그리드
    │  • 1~5자루 → 30분 차단 (1슬롯 점유)
    │  • 6자루+ → 60분 차단 (2슬롯 점유)
    │  • 충돌 검사: consultations 매장방문/출장 ↔ repairs 직접방문 양방향
  → 방문 시간 선택 → submit (proceed_type='직접방문', visit_date/time)
  → /api/repair/public/submit
    • visit_duration_min 서버 계산 (qty + thresholdQty)
    • address 컬럼 NULL (매장 워크인, 주소 불필요)
    • shipping_fee = 0
    • Google Calendar 이벤트 자동 생성 (lib/google/repair-calendar-sync.ts, colorId '6' Tangerine)
  → 알림톡 (as_visit_booked) — Phase 4 검수 통과 후 발송
  → [status: intake]
  → (관리자) 접수확인 → confirmed_at
  → 방문 D-1 21:00 + 당일 09:00 리마인드 (as_visit_remind) — cron 신규, Phase 4
  → 매장 방문 → 즉석 검수 + 수리
  → 사이드 패널 "고객 수령 완료" 버튼
  → [status: delivered] (배송 없으니 delivered=수령완료 의미 전환)
    → 알림톡 (as_visit_completed) — Phase 4 검수 통과 후 발송 (후기 링크 포함)
  → [status: completed]
  ※ 시간 변경: TMS PATCH visit_date/time → as_visit_rescheduled + Google 이벤트 자동 update
  ※ 취소: status='cancelled' → as_visit_cancelled + Google 이벤트 자동 삭제
```

**직접방문 핵심 설계 원칙**:
- 입출고 단계 스킵 (현장 즉시 검수, 송장 없음)
- 결제는 매장 사후 결제 (검수 후 금액 확정)
- DB 분기 컬럼 3개만 추가 (`visit_date`/`visit_time`/`visit_duration_min`) — 092 마이그레이션
- 슬롯 정책 4컬럼 추가 (`repair_slot_step_min`/`threshold_qty`/`block_under_min`/`block_over_min`) — 092
- Google Calendar 동기화 컬럼 2개 추가 (`google_event_id`/`updated_at`) — 093
- 기존 방문수거/직접발송 흐름 무영향 (proceed_type 분기로 격리)

### 🆕 판매건 합포장 출고 분기 (2026-05-24 추가)

**케이스**: 같은 고객이 같은 시기 제품 주문도 하여, **판매건 송장에 복원수리를 합쳐 발송**한 경우. 복원수리 단독 송장 없음.

```
출고대기 [ready_to_ship, invoice_number=NULL]
  → (관리자) 사이드 패널 "판매건 합포장 출고" 버튼 클릭
  → MergedShipModal 자동 열림
    → GET /api/repair/[id]/related-shipments
      → 같은 phone(정규화)으로 offline_sales + orders 검색
      → 송장 보유 + 미취소/배송중·완료 건만 노출
  → (관리자) 검색 결과 1클릭 선택 (또는 fallback 직접 송장번호 입력)
  → 알림톡 발송 체크박스 (default ON — 수리내역서 노출로 신뢰감)
  → POST /api/repair/[id]/merged-ship
    → 선택한 송장의 invoice_number/courier_name → repairs 에 복사
    → status='shipped' + shipped_at=now
    → repair_history 이력 박제 (출처: sale/order/manual + id)
    → 알림톡 (as_shipped) 발송 — invoice_number 채워짐 → 기존 흐름 그대로
      → [수리내역 조회] 버튼 정상 (#{as_uid})
      → [배송조회] 버튼 정상 (#{tracking_number} = 판매건 송장)
  → [status: shipped] (출고완료) — 일반 출고와 동일 카드로 표시
```

**핵심 설계 원칙**:
- 새 솔라피 템플릿 0개 (기존 `as_shipped` 그대로 활용)
- 새 DB 컬럼 0개 (invoice_number 단순 복사로 모든 흐름 자동 작동)
- 매출 집계 영향 0 (RPC 088 status != 'cancelled' 조건 — 이미 잡혀있음)
- 자동 매칭 효율: 같은 phone 정규화 → 사장님 검색 수고 0

### 상태 전이 규칙
```
intake → [pickup_scheduled, cost_notified, cancelled]
pickup_scheduled → [cost_notified, cancelled]
cost_notified → [repairing, cancelled]
repairing → [ready_to_ship, cancelled]
ready_to_ship → [shipped]  // (1) 송장 생성 → 출고완료 (정상)
                          // (2) 판매건 합포장 출고 → POST /merged-ship (송장 복사)
shipped → [delivered]      // 🆕 (2026-05-24) ALPS 인수자등록('91') 자동 감지 → cron 자동 전환
                          //   (4시간마다 폴링, /api/cron/track-delivery)
                          //   fallback: 사이드 패널 "수동 배송완료 처리" 텍스트 링크
shipped → [delivered]
delivered → [completed]
completed → (terminal)
cancelled → (terminal)

* paid_at: 파이프라인과 독립된 플래그 (어느 상태에서든 입금확인 가능)
* proceed_type 필터:
  - 방문수거: pickup_scheduled 필수
  - 직접발송: pickup_scheduled 생략
  - 🆕 직접방문(당일수리): pickup_scheduled / cost_notified / paid_at 모두 스킵 가능
    intake → (즉석 검수+수리+수령) → delivered → completed (단축 흐름)
    visit_date/visit_time 필수, address NULL, shipping_fee 0
```

---

## 2. 시스템 연동 흐름

```
[고객] ──폼──→ [Vercel API: /api/repair/public/submit]
                  │
         ┌───────┼───────────┐
         ▼       ▼           ▼
   [Google 시트] [Make]   [TMS /api/repair/sync]
                  │           │
                  ▼           ▼
             [Solapi]    [Supabase DB]
             (알림톡)         │
                              ▼
                        [TMS React UI]
                              │
            ┌─────────────────┼────────────────┐
            ▼                 ▼                ▼
   [검수/비용안내]     [송장생성(ALPS)]    [알림톡 발송]
   PATCH /repair/[id]  POST /repair/[id]/ship  POST /repair/[id]/notify
                              │                     또는 PATCH 자동
                              ▼
                        [롯데택배 ALPS API]
                        (운송장 12자리 생성)
```

### 알림톡 발송 구조 (4가지 트리거)

| 템플릿 | 트리거 | 경로 |
|--------|--------|------|
| `as_received` (접수완료) | Vercel submit route | Vercel → Make webhook |
| `as_cost_notice` (비용안내) | UI "비용안내" 버튼 | UI → POST `/api/repair/[id]/notify` |
| `as_payment_confirmed` (입금확인) | paid_at 플래그 설정 시 | PATCH `/api/repair/[id]` → after() 자동 |
| `as_shipped` (출고완료) | shipped 상태 전환 시 | PATCH `/api/repair/[id]` → after() 자동 |
| `as_cancelled` (취소안내) | cancelled 상태 전환 시 | PATCH `/api/repair/[id]` → after() 자동 |
| `as_satisfaction` (만족도) | 미구현 | — |

### Make 웹훅 URL 분기 (2026-03-05)
| 웹훅 | 환경변수 | 대상 템플릿 |
|------|----------|-------------|
| 상담 알림톡 | `MAKE_WEBHOOK_URL` | confirmed, cancelled, suggest 등 상담 17종 |
| 복원수리 접수/취소 | Vercel API | as_received (Vercel → Make webhook) |
| 복원수리 상태변경 | `MAKE_REPAIR_WEBHOOK_URL` | as_cost_notice, as_payment_confirmed, as_shipped, as_cancelled, as_satisfaction |

### 웹훅 payload 주요 필드
| 필드 | 값 | 용도 |
|------|-----|------|
| `id` | repair.as_id (AS-YYYYMMDD-NNN) | Make `#{as_uid}` → 수리내역서 링크 |
| `tracking` | 송장번호 | 배송조회 |
| `courier` | 롯데택배 (고정) | 배송조회 버튼 활성화 |
| `as_amount` / `shipping_amount` / `total_amount` | 금액 | 비용 안내 |
| `pickup_date` | `YYYY년 MM월 DD일 (X요일)` 포맷 | 방문수거 접수 알림톡 수거예정일 (as_received) · 직접발송은 빈 문자열 |
| `delivery_method` | 문앞보관 / 매장카운터 / 직접전달 | 방문수거 수령 방식 (as_received) |

### 비용 계산 규칙
```
서비스비 = (마모루 수량 × 10,000) + (타사 수량 × 20,000)
배송비   = 방문수거일 때: 1개=6,000 / 2개=3,000 / 3개+=무료
           직접발송일 때: 1개=3,000 / 2개+=무료
합계     = 서비스비 + 배송비
※ 실제 우체국 수거+발송 비용 8,000원 대비 할인 적용
```

---

## 3. 구현 완료 ✅

### API Routes (7개)
| 엔드포인트 | 메서드 | 기능 |
|------------|--------|------|
| `/api/repair` | GET | 목록 조회 (상태그룹 필터/검색/페이징) |
| `/api/repair` | POST | 신규 접수 생성 |
| `/api/repair/[id]` | GET | 단건 + 검수 + 이력 조회 |
| `/api/repair/[id]` | PATCH | 상태/필드 변경 + 이력 + 자동 알림톡 |
| `/api/repair/[id]/inspect` | POST/PUT | 검수 데이터 저장/수정 |
| `/api/repair/[id]/notify` | POST | 수동 알림톡 발송 |
| `/api/repair/[id]/ship` | POST/DELETE | 송장 생성(ALPS) / 취소 |
| `/api/repair/[id]/related-shipments` | GET | 🆕 같은 phone 의 offline_sales + orders 송장 보유건 검색 (합포장 매칭용, 2026-05-24) |
| `/api/repair/[id]/merged-ship` | POST | 🆕 판매건 합포장 출고 처리 — invoice 복사 + status='shipped' + as_shipped 자동 발송 (2026-05-24) |
| `/api/repair/sync` | POST | 외부 동기화 (레거시, 사용 빈도 낮음) |
| `/api/repair/report` | GET | 수리내역 공개 API (CORS, 인증 불필요) |

### 탭별 필터 조건 (use-repair-tabs.ts) — 2026-05-19 정정

| 탭 | 필터 조건 |
|---|---|
| **신규접수** | `status='intake' AND confirmed_at IS NULL` (접수확인 전 — 방문수거·직접발송 모두) |
| **수거접수필요** | `status='intake' AND proceed_type='방문수거' AND confirmed_at IS NOT NULL` ⭐ (접수확인 완료한 방문수거 = 수거 예약 필요) |
| **입고대기** | (직접발송) `status='intake' AND proceed_type!='방문수거' AND confirmed_at IS NOT NULL` + (수거완료) `status='pickup_scheduled'` |
| **진행중** | `status IN ('cost_notified', 'repairing')` |
| **출고대기** | `status='ready_to_ship'` |
| **출고완료** | `status IN ('shipped', 'delivered', 'completed')` |

> 🐛 **2026-05-19 fix**: "수거접수필요" 탭이 `confirmed_at IS NULL`(접수확인 전)로 잘못 설정되어 있었음. 방문수거 건에 "접수확인" 누르면 → confirmed_at 채워짐 → 신규접수·수거접수필요·입고대기 3개 탭 조건 모두 탈락 → **orphan(리스트에서 사라짐)**. 대시보드 카운트는 status 기반이라 별개로 정상 표시되어 불일치 발견. `confirmed_at IS NOT NULL`로 정정 → 접수확인 후 수거접수필요 탭에 정상 표시.

### 정상 흐름 (정정 후)
```
신규접수(미확인 전체)
   │ 접수확인 (confirmed_at 기록)
   ├─[방문수거] → 수거접수필요 → 수거접수완료(pickup_scheduled) → 입고대기
   └─[직접발송] → 입고대기(intake+confirmed)
```

### 컴포넌트 (15+개)
- 고정 탭 바 6개: 신규접수 / 수거접수필요 / 입고대기 / 진행중 / 출고대기 / 출고완료
- **인라인 퀵 액션** (03-26 리모델): 탭별 특화 — 접수확인/수거접수/입금확인/송장생성/출고완료
- **상단 요약 카드 4개** (03-26→03-30): 신규접수(blue) / 진행중(amber) / 미입금(red) / 3일경과(orange) — **클릭 시 탭 이동 + 교차 필터 활성** (토글)
- **useRepairTabData 기반 리스트** (03-26): 7탭→6탭 전환, 비즈니스 로직 쿼리 활용
- **isLg conditional rendering** (03-26): CSS hidden→조건부 렌더링 (SlidePanel 이슈 해결)
- PC 마스터-디테일 레이아웃 (lg+) — 좌측 480px 고정 + 우측 flex-1 상세 모니터
- 검수 폼 (가위별 7항목: 날끝/중간/안쪽/빗살/텐션/부품/스토퍼)
- 검수 요약 (읽기 전용 테이블)
- 사이드바 액션 카드 (비용+상태버튼+입금/송장)
- 타임라인 (상태 변경 이력)
- 상태 배지 (색상 코딩 + 진행방식 배지)
- **탭별 특화 정보 칩** (03-26): 수거필요→주소, 진행중→입금상태, 출고대기→송장/포장, 출고완료→배송상태
- **완료/취소 시각 구분** (03-26): green border-left / opacity+취소선 / 미입금=orange border / 미입금+3일경과=red border
- **입금확인 알림톡 선택** (03-30): 체크박스로 발송/미발송 분리 (skip_notify 플래그)
- **확인 모달 전체 적용** (03-29): 비용안내/입금확인/출고완료/송장취소/수리취소 — ConfirmModal 공통 컴포넌트

### 검수 자동 문구 생성
- `lib/repair/inspection-text.ts` — 검수 데이터 분석 → 한국어 작업 설명 자동 생성
- 무뎌짐, 찍힘, 빗살 손상, 텐션 느슨, 부품 교체, 스토퍼 교체 자동 감지

### DB 테이블
- `repairs` — 메인 (as_id, 비용, 상태, 송장, confirmed_at, packed_at, paid_at)
- `repair_inspections` — 가위별 검수 (blade_tip/mid/inner, comb, tension, parts, stopper)
- `repair_history` — 상태 변경 이력

### ~~GAS 스크립트~~ (제거 완료 — 2026-04-01)
- ~~`projects/as/Code.gs`~~ — Vercel API로 완전 대체, 보관처리 예정

### 고객 대면 페이지 (GitHub Pages)
- `page_form.html` — 통합 접수 폼 (마모루+타사)
- `page_guide.html` — 복원수리 안내 (아임웹 iframe용, 6탭: 마모루복원수리/과정안내/소요시간/비용안내/포장방법/QnA, 라이트모드+다크 intro, CTA 통합접수)
- `page_as_guide.html` — 복원수리 안내 (알림톡 링크용, 모바일 전용 4탭)
- `page_as_report.html` — 수리내역 조회 (TMS API 연동)

---

## 4. 완료된 외부 연동 ✅ (2026-03-03)

| 항목 | 완료일 |
|------|--------|
| 솔라피 복원수리 5종 템플릿 등록 + 검수 승인 | 2026-03-03 |
| GAS Script Properties 설정 (TMS_REPAIR_SYNC_URL, CRON_SECRET) | 2026-02-28 |
| 복원수리 접수 → GAS → TMS 동기화 테스트 통과 | 2026-02-28 |

## 5. 미완료 ❌

| 항목 | 의존성 | 우선순위 |
|------|--------|----------|
| ~~Make Router에 복원수리 6종 분기 추가~~ | ~~완료~~ | ~~✅ 03-12~~ |
| ~~솔라피 as_cancelled 취소안내 템플릿~~ | ~~완료~~ | ~~✅ 03-03~~ |
| ~~BC 메타데이터 제거 후 솔라피 재검수~~ | ~~완료~~ | ~~✅ 03-19~~ |
| ~~사진 업로드 Supabase Storage 연동~~ | ~~완료~~ | ~~✅ 03-24~~ |
| 복원수리 E2E 전체 플로우 검증 | Make 분기 + 솔라피 재검수 완료 후 | 높음 |
| 만족도 알림톡 (as_satisfaction) 자동 발송 로직 | 솔라피 검수 | 중간 |
| 주소 수정 시 다음 주소검색 API 연동 | 없음 | 중간 |
| 사진 마킹 (photo-marker) html2canvas | 없음 | 낮음 |
| 수리내역서 자동 생성 (Before/After 웹카드) | 사진 업로드 ✅ | 낮음 |

---

## 6. 핵심 파일 맵

### TMS API
| 파일 | 설명 |
|------|------|
| `app/src/app/api/repair/route.ts` | GET/POST 목록/생성 |
| `app/src/app/api/repair/[id]/route.ts` | GET/PATCH 단건 + 상태머신 |
| `app/src/app/api/repair/[id]/inspect/route.ts` | POST/PUT 검수 데이터 |
| `app/src/app/api/repair/[id]/notify/route.ts` | POST 알림톡 발송 |
| `app/src/app/api/repair/[id]/ship/route.ts` | POST/DELETE 송장 |
| `app/src/app/api/repair/sync/route.ts` | POST GAS 동기화 |
| `app/src/app/api/repair/report/route.ts` | GET 공개 수리내역 API |

### TMS UI
| 파일 | 설명 |
|------|------|
| `app/src/app/(dashboard)/repairs/page.tsx` | 복원수리 메인 (탭 바+목록) |
| `app/src/app/(dashboard)/repairs/dashboard/page.tsx` | 복원수리 대시보드 |
| `app/src/app/(dashboard)/repairs/[id]/page.tsx` | 복원수리 상세 (모바일) |
| `app/src/components/repairs/repair-tab-bar.tsx` | 고정 탭 바 6개 |
| `app/src/components/repairs/repair-list.tsx` | 마스터 목록 |
| `app/src/components/repairs/repair-detail-panel.tsx` | PC 상세 패널 |
| `app/src/components/repairs/repair-detail-card.tsx` | 고객/접수 정보 카드 |
| `app/src/components/repairs/repair-action-chips.tsx` | 인라인 액션 칩 |
| `app/src/components/repairs/inspection-form.tsx` | 검수 입력 폼 |
| `app/src/components/repairs/inspection-summary.tsx` | 검수 요약 |
| `app/src/components/repairs/repair-timeline.tsx` | 이력 타임라인 |
| `app/src/components/repairs/sidebar-action-card.tsx` | 사이드바 액션 |
| `app/src/components/repairs/repair-status-badge.tsx` | 상태 배지 |
| `app/src/components/repairs/tabs/*.tsx` | 탭별 구현 (6개) |

### TMS Lib
| 파일 | 설명 |
|------|------|
| `app/src/lib/repair/transitions.ts` | 상태 전이 규칙 (v3) |
| `app/src/lib/repair/cost-calculator.ts` | 비용 계산 |
| `app/src/lib/repair/inspection-text.ts` | 검수→작업문구 자동생성 |
| `app/src/lib/repair/sync.ts` | GAS→DB 동기화 로직 |
| `app/src/lib/notification/make-webhook.ts` | Make webhook (상담과 공유) |
| `app/src/hooks/use-repairs.ts` | React Query 훅 8개 |
| `app/src/hooks/use-repair-tabs.ts` | 탭별 쿼리+카운트 훅 |

### GAS (Google Apps Script)
| 파일 | 설명 |
|------|------|
| `projects/as/Code.gs` | 접수/관리/ALPS/TMS동기화 |

### 롯데택배 ALPS
| 파일 | 설명 |
|------|------|
| `app/src/lib/lotte/client.ts` | ALPS API (book/cancel/track) |
| `app/src/lib/lotte/types.ts` | 타입 정의 |

### DB 마이그레이션
| 파일 | 설명 |
|------|------|
| `app/supabase/migrations/005_repair_paid_at.sql` | paid_at 플래그 |
| `app/supabase/migrations/006_repair_confirmed_packed.sql` | confirmed_at, packed_at |

### 고객 대면 페이지
| 파일 | 설명 |
|------|------|
| `projects/as/page_form.html` | 통합 접수 폼 |
| `projects/as/page_as_report.html` | 수리내역 조회 |
