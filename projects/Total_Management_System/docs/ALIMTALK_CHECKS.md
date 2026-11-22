# 알림톡 종류별 점검 — 템플릿 칸 · 읽는 값 · 보내는 곳 (2026-11-22)

> 작업실 할 일 `t-261122-alimtalk-checks`. 계기: 출장 접수 알림톡에 희망 요일 · 시간대가 빈칸으로 나간 일(`t-261121-field-wish-alimtalk`).

## 어떻게 점검하나
- **Make 쪽 표**: `app/tests/alimtalk/make-reads.json` — 시나리오 7개 · 분기 48개가 템플릿 칸을 어떤 값에서 읽는지. Make 를 고치면 `node tests/alimtalk/snapshot-make.mjs` 로 다시 뽑는다(토큰은 저장소 밖에서 읽음).
- **TMS 쪽**: `app/tests/alimtalk/scan.ts` 가 `sendNotification(...)` 을 부르는 곳을 전부 글자로 읽어, 신호마다 보내는 칸을 뽑는다. 실제로 보내지 않는다.
- **점검**: `app/tests/alimtalk/vars.test.ts` — `npm run check` 에 같이 돈다. 템플릿 칸이 읽는 값을 TMS 가 안 보내면 실패한다. 알고 있는 빈칸은 `KNOWN_GAPS` 에 이유와 함께 둔다(고치면 지운다 — 안 지우면 그것도 실패).
- **못 잡는 것**: 값 이름은 맞는데 실제 값이 비는 경우(예: 고객이 고르지 않은 칸)는 흐름 점검(`tests/flows/consult-alimtalk-vars.test.ts` 등)이 맡는다. 솔라피 템플릿 본문 자체의 오타는 Make 표에 안 보이므로 못 잡는다.

## 2026-11-22 첫 점검에서 나온 것
1. **복원수리 출고(`as_shipped`)** — 템플릿 칸 `#{tracking_number}` 를 Make 가 `tracking_number` 에서 읽는데 TMS 는 `tracking` 으로 보낸다 → 송장번호 빈칸 가능 (4곳: 송장 자동 감지 크론 · 합포장 · 상태 변경 · 함께 출고).
2. **매장방문 확정 재발송(`confirmed`, 상담 상세 [알림톡 재발송])** — 일정 변경 버튼 주소(`change_request_link`)를 안 보낸다.
3. **EVENT 입금 안내(`event_payment_notice`)** — TMS 는 보내는데 Make 에 분기가 없어 아무것도 안 나간다.
4. **복원수리 직접 전달(`as_handed_over`)** — 템플릿 등록 전이라 Make 분기 없음(작업실 `t-261121-handover-template`, 알고 있는 것).

## 신호별 표

### 05 아임웹 클레임 (취소·반품·교환)

| 신호 | Make 분기 | 템플릿 칸 ← 읽는 값 | 보내는 곳(TMS) |
|---|---|---|---|
| `imweb_cancel_requested` | IW-주문취소접수 | name ← name<br>order_no ← order_no<br>product_name ← product_name | `lib/imweb/claim-notify.ts` |
| `imweb_cancel_completed` | IW-취소처리완료 | order_no ← order_no<br>product_name ← product_name | `lib/imweb/claim-notify.ts` |
| `imweb_return_requested` | IW-반품요청접수 | name ← name<br>order_no ← order_no<br>product_name ← product_name | `lib/imweb/claim-notify.ts` |
| `imweb_return_approved` | IW-반품승인 | order_no ← order_no<br>product_name ← product_name | `lib/imweb/claim-notify.ts` |
| `imweb_return_completed` | IW-반품완료 | order_no ← order_no<br>product_name ← product_name | `lib/imweb/claim-notify.ts` |
| `imweb_exchange_requested` | IW-교환접수 | name ← name<br>order_no ← order_no<br>product_name ← product_name | `lib/imweb/claim-notify.ts` |

### 01 상담 (매장방문·출장·톡상담)

| 신호 | Make 분기 | 템플릿 칸 ← 읽는 값 | 보내는 곳(TMS) |
|---|---|---|---|
| `confirmed` | 매장방문_확정 | name ← name<br>visit_date ← visit_date<br>visit_time ← visit_time<br>change_request_link ← change_request_link | `app/api/consultation/admin-create/route.ts`<br>`app/api/consultation/notify/route.ts`<br>`app/api/consultation/public/submit/route.ts`<br>`app/api/consultation/[id]/route.ts` |
| `rescheduled` | 매장방문_일정변경 | new_date ← new_date<br>new_time ← new_time<br>change_request_link ← change_request_link | `app/api/consultation/[id]/route.ts` |
| `remind24` | 매장방문_RE24 | visit_date ← visit_date<br>visit_time ← visit_time | `app/api/cron/send-reminders/route.ts` |
| `remind2` | 매장방문_RE2 | visit_time ← visit_time | `app/api/cron/send-reminders/route.ts` |
| `cancelled` | 매장방문_취소 | (칸 없음) | `app/api/consultation/public/cancel/route.ts`<br>`app/api/consultation/[id]/route.ts` |
| `request` | 출장_접수완료 | name ← name<br>address ← address<br>visit_date ← hope_days_text<br>visit_time ← hope_times_text | `app/api/consultation/public/submit/route.ts` |
| `suggest` | 출장_시간제안 | name ← name<br>confirm_link ← confirm_link | `app/api/consultation/suggest/route.ts` |
| `field_confirmed` | 출장_확정 | address ← address<br>visit_date ← visit_date<br>visit_time ← visit_time<br>change_request_link ← change_request_link | `app/api/consultation/admin-create/route.ts`<br>`app/api/consultation/public/confirm/route.ts`<br>`app/api/consultation/[id]/route.ts` |
| `field_rescheduled` | 출장_일정변경 | address ← address<br>visit_date ← visit_date<br>visit_time ← visit_time<br>change_request_link ← change_request_link | `app/api/consultation/[id]/route.ts` |
| `field_remind_24h` | 출장_RE24 | address ← address<br>visit_date ← visit_date<br>visit_time ← visit_time | `app/api/cron/send-reminders/route.ts` |
| `field_remind_2h` | 출장_RE2 | address ← address<br>visit_time ← visit_time | `app/api/cron/send-reminders/route.ts` |
| `field_cancelled` | 출장_취소 | address ← address<br>visit_date ← visit_date<br>visit_time ← visit_time | `app/api/consultation/public/cancel/route.ts`<br>`app/api/consultation/[id]/route.ts` |
| `field_delayed` | 출장_도착지연 | delay_min ← delay_min<br>visit_time ← visit_time<br>visit_time_revised ← visit_time_revised | `app/api/consultation/delay/route.ts` |
| `talk_received` | 톡상담_접수 | name ← name | `app/api/consultation/public/submit/route.ts` |
| `talk_ready` | 톡상담_진행 | (칸 없음) | `app/api/consultation/[id]/route.ts` |
| `review_request` | 상담_후기요청 | name ← name<br>phone ← phone<br>subtype ← subtype<br>consult_uid ← consult_uid | `lib/notification/review-request.ts` |

### 03 판매·배송 (아임웹·오프라인 공통)

| 신호 | Make 분기 | 템플릿 칸 ← 읽는 값 | 보내는 곳(TMS) |
|---|---|---|---|
| `sales_shipped` | 제품_출고완료 | name ← name<br>courier ← courier<br>tracking ← tracking<br>goods_name ← goods_name | `lib/notification/sales-shipped.ts` |
| `purchase_review_request` | 제품구매_후기요청 | name ← name<br>order_uid ← order_uid | `app/api/imweb/orders/[id]/review-request/route.ts`<br>`app/api/orders/[id]/pickup-complete/route.ts`<br>`lib/imweb/sync.ts`<br>`lib/notification/review-request.ts` |

### 04 반품·교환 (오프라인 판매분)

| 신호 | Make 분기 | 템플릿 칸 ← 읽는 값 | 보내는 곳(TMS) |
|---|---|---|---|
| `return_received` | 반품교환_회수예약 | name ← name<br>return_type ← return_type<br>product_name ← product_name<br>return_number ← return_number | `app/api/returns/route.ts` |
| `return_inbound` | 반품교환_입고완료 | product_name ← product_name<br>return_number ← return_number | `app/api/returns/[id]/route.ts` |
| `exchange_shipped` | 교환_발송완료 | courier ← courier<br>tracking ← tracking<br>product_name ← product_name | `app/api/orders/[id]/exchange-ship/route.ts`<br>`app/api/returns/[id]/ship/route.ts` |

### 06 EVENT

| 신호 | Make 분기 | 템플릿 칸 ← 읽는 값 | 보내는 곳(TMS) |
|---|---|---|---|
| `event_received` | EVENT_접수완료 | name ← name<br>items ← items<br>event_name ← event_name<br>event_detail ← event_detail | `app/api/event/public/submit/route.ts` |
| `event_payment_confirmed` | EVENT_입금확인 | event_name ← event_name<br>event_number ← event_number<br>total_amount ← total_amount | `app/api/events/[id]/route.ts` |
| `event_shipped` | EVENT_출고완료 | id ← id<br>courier ← courier<br>tracking ← tracking<br>event_name ← event_name<br>goods_name ← goods_name | `lib/notification/sales-shipped.ts` |
| `review_event_won` | 리뷰이벤트_당첨안내 | name ← name<br>rank ← rank<br>prize ← prize<br>token ← token | `lib/reviews/event-shipments.ts` |
| `review_event_shipped` | 리뷰이벤트_출고안내 | name ← name<br>prize ← prize<br>courier ← courier<br>tracking ← tracking | `lib/reviews/event-shipments.ts` |

### 02 복원수리 (접수·상태변경·후기)

| 신호 | Make 분기 | 템플릿 칸 ← 읽는 값 | 보내는 곳(TMS) |
|---|---|---|---|
| `as_received` | 직접발송 | name ← name<br>address ← address<br>delivery_method ← delivery_method | `lib/repair/intake.ts` |
| `as_received` | 방문수거_문앞 | name ← name<br>address ← address<br>pickup_date ← pickup_date<br>delivery_method ← delivery_method | `lib/repair/intake.ts` |
| `as_received` | 방문수거_카운터 | name ← name<br>address ← address<br>pickup_date ← pickup_date<br>delivery_method ← delivery_method | `lib/repair/intake.ts` |
| `as_received` | 방문수거_직접전달 | name ← name<br>address ← address<br>pickup_date ← pickup_date<br>delivery_method ← delivery_method | `lib/repair/intake.ts` |
| `as_visit_booked` | 직접방문_접수완료 | qty ← qty<br>name ← name<br>as_id ← as_id<br>visit_date ← visit_date<br>visit_time ← visit_time<br>visit_duration_min ← visit_duration_min<br>change_request_link ← change_request_link | `lib/repair/intake.ts` |
| `as_cost_notice` | 입고&비용안내 | as_amount ← as_amount<br>qty_other ← qty_other<br>qty_mamoru ← qty_mamoru<br>total_amount ← total_amount<br>shipping_amount ← shipping_amount | `app/api/repair/[id]/notify/route.ts` |
| `as_payment_confirmed` | 입금확인완료 | (칸 없음) | `lib/repair/payment-notify.ts` |
| `as_shipped` | 출고완료&송장 | as_uid ← as_uid<br>courier ← courier<br>tracking_number ← tracking_number | `app/api/cron/track-delivery/route.ts`<br>`app/api/repair/[id]/merged-ship/route.ts`<br>`app/api/repair/[id]/route.ts`<br>`lib/notification/combined-shipped.ts` |
| `as_cancelled` | 수리_취소 | (칸 없음) | `app/api/repair/[id]/route.ts` |
| `as_review_request` | 수리_후기요청 | name ← name<br>as_uid ← as_uid | `app/api/repair/[id]/route.ts`<br>`lib/notification/review-request.ts` |
| `as_visit_remind_24h` | 직접방문_RE24 | qty ← qty<br>visit_date ← visit_date<br>visit_time ← visit_time<br>change_request_link ← change_request_link | `app/api/cron/repair-visit-remind/route.ts` |
| `as_visit_remind_2h` | 직접방문_RE2 | qty ← qty<br>visit_time ← visit_time | `app/api/cron/repair-visit-remind/route.ts` |
| `as_visit_rescheduled` | 직접방문_일정변경 | qty ← qty<br>visit_date ← visit_date<br>visit_time ← visit_time<br>change_request_link ← change_request_link | `app/api/repair/public/resched/route.ts` |
| `as_visit_cancelled` | 직접방문_취소 | visit_date ← visit_date<br>visit_time ← visit_time | `app/api/repair/public/cancel/route.ts`<br>`app/api/repair/[id]/route.ts` |

### 07 출장 맡김·함께 출고

| 신호 | Make 분기 | 템플릿 칸 ← 읽는 값 | 보내는 곳(TMS) |
|---|---|---|---|
| `as_field_received` | TMS-출장맡김확인 | qty ← qty<br>name ← name<br>as_id ← as_id<br>as_uid ← as_uid<br>pay_status ← pay_status | `lib/repair/intake.ts` |
| `combined_shipped` | TMS-함께출고 | qty ← qty<br>name ← name<br>as_id ← as_id<br>as_uid ← as_uid<br>courier ← courier<br>tracking ← tracking<br>product_name ← product_name | `lib/notification/combined-shipped.ts` |