# TMS 흐름도 — EVENT (고객 접수 이벤트)

> 최종 업데이트: 2026-09-13 · 알림톡 범용화(이벤트명·고객 안내 문구 변수) / (최초) 2026-06-15 재고 전환 이벤트 1탄으로 구축. 접수형 이벤트 공통 허브.

## 개요
릴스/DM → **EVENT 접수 페이지(카탈로그형)** → 접수 → 입금안내 → 입금확인 → **판매(offline_sales) 자동 전환** → 발송/배송완료/후기(기존 판매 인프라). 캠페인(이벤트)별로 분리 관리.

## 데이터 모델
- `event_campaigns` (마이그 104·105): 이벤트 단위. `type`(stock_clearance/limited/group_buy/tester/trade_in/other), `status`(active/ended), `is_default`, `discount_rules` jsonb, `customer_notice`(마이그 145 — 신청완료 알림톡 #{event_notice}).
- `event_submissions` (마이그 103·104): 접수 1건. `event_number`(EV-YYYYMMDD-NNN), `campaign_id`, `items` jsonb, `slicing_addon`, `total_amount`, `status`, `payment_noticed_at`, `paid_at`, `sale_id`(전환 시), `receive_method`(delivery/visit).
- `event_history`: 상태 이력.

## 상태 파이프라인
```
received(접수)
  └─[입금안내]→ payment_noticed(입금대기)   ← 사장님이 재고확인 후, 입금안내 알림톡
        └─[입금확인]→ converted(판매전환)    ← 입금확인 알림톡 + offline_sales 자동생성·재고차감
  └─[취소]→ cancelled
```
전환 후(=offline_sales): 송장생성→**출고완료 알림톡(event_shipped, EVENT 전용)** / ALPS cron→배송완료 자동 / 약속✓→후기요청 자동. (출고 판별 = 판매 memo `EVENT 전환` 접두어 → sales-shipped.ts 에서 event_shipped 분기, 나머지 인프라 재사용)

## 가격 (lib/event/pricing.ts = 서버 권위 / page_form.html JS 동일 복제)
- 품목 단가×수량 + 슬라이싱 가공(+20,000/자루).
- **묶음 할인**: 캠페인 `discount_rules[{unit_price,min_qty,bundle_price}]`. **같은 단가끼리**, min_qty 도달 시 **묶음 반복 + 나머지 정가**. 혼합 미적용.
  - 예 50000/3/130000 → 3자루=13만, 4자루=18만, 6자루=26만.
- 판매 전환 시 묶음할인 = `offline_sales.discount_amount`.

## 알림톡 (make-webhook.ts, **webhook_event 전용 시나리오** — 2026-07-31 분리)
- `event_received`(접수확인+비용안내·자동) / `event_payment_confirmed`(입금확인·자동) / **`event_shipped`(출고완료·자동, EVENT 전용 신규)** (+선택 `event_payment_notice`).
- 4개 웹훅 = consultation/as_received/repair/**event(`MAKE_EVENT_WEBHOOK_URL`, 설정 `notifications.webhook_event`, 미설정 시 consultation 폴백)**.
- 출고완료는 `sales_shipped` 대신 `event_shipped` — sales-shipped.ts 가 memo `EVENT 전환` 접두어면 자동 분기(집하 cron·수동 [출고완료] 공통).
- 셋업·전환순서 = `docs/EVENT_ALIMTALK_SETUP.md`. ※ 솔라피 콘솔 3종 검수 + Make 분기 필요.

### 2026-09-13 범용 템플릿 변수 (특정 이벤트에 묶이지 않게)
- `lib/event/campaign-notify.ts` 공용 헬퍼: `event_name` = 캠페인명, `event_notice` = 캠페인 `customer_notice` (비면 "추가로 필요한 사항이 있으면 따로 연락드립니다"). 캠페인 없으면 이벤트명 "MAMORU 이벤트". select('*') 라 마이그 145 전에도 안 깨짐.
- `event_received`(접수) → `event_name`·`event_notice` 추가 / `event_payment_notice`·`event_payment_confirmed` → `event_name` 추가(재고판매 LS 제외) / `event_shipped` → 판매 sale_id → event_submissions → 캠페인으로 `event_name` 추가.
- 기존 변수는 모두 유지(추가만) → 이전 템플릿도 그대로 동작.
- 캠페인 설정 모달에 "신청완료 알림톡 안내 문구"(80자) 입력. 마이그 전 저장 시 409 안내.
- 이름 규칙: 신청완료만 `#{name}`, 입금확인·출고완료(후속 안내)는 이름 없음. 템플릿 등록본 = `docs/EVENT_ALIMTALK_SETUP.md` §9.

## 화면/코드
- 고객폼: `projects/event/page_form.html` (page.mamoru.kr, `?campaign=<id>`). 공개 API `app/api/event/public/{products,submit}`.
- 허브: `app/(dashboard)/events/page.tsx` (캠페인 카드 → 접수목록). hooks `use-events.ts`.
- 캠페인 API: `app/api/campaigns/route.ts`(GET/POST) + `[id]/route.ts`(PATCH). 접수 API `app/api/events/...`.
- 전환: `lib/event/convert-to-sale.ts`.
- 매장 워크인(사전접수 X) = `/sales/new`에서 EVENT 품목(category='EVENT') 선택 = 일반 판매. 특수처리 없음.

## 미래 (방향)
접수형(한정/공동구매/체험단/트레이드인)=캠페인 추가. 멤버십(서포터즈/앰버서더/테스터)=고객 등급(price_groups) 재사용, 별도 사이클.
