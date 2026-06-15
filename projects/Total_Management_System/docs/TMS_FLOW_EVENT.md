# TMS 흐름도 — EVENT (고객 접수 이벤트)

> 최종 업데이트: 2026-06-15 · 재고 전환 이벤트 1탄으로 구축. 접수형 이벤트 공통 허브.

## 개요
릴스/DM → **EVENT 접수 페이지(카탈로그형)** → 접수 → 입금안내 → 입금확인 → **판매(offline_sales) 자동 전환** → 발송/배송완료/후기(기존 판매 인프라). 캠페인(이벤트)별로 분리 관리.

## 데이터 모델
- `event_campaigns` (마이그 104·105): 이벤트 단위. `type`(stock_clearance/limited/group_buy/tester/trade_in/other), `status`(active/ended), `is_default`, `discount_rules` jsonb.
- `event_submissions` (마이그 103·104): 접수 1건. `event_number`(EV-YYYYMMDD-NNN), `campaign_id`, `items` jsonb, `slicing_addon`, `total_amount`, `status`, `payment_noticed_at`, `paid_at`, `sale_id`(전환 시), `receive_method`(delivery/visit).
- `event_history`: 상태 이력.

## 상태 파이프라인
```
received(접수)
  └─[입금안내]→ payment_noticed(입금대기)   ← 사장님이 재고확인 후, 입금안내 알림톡
        └─[입금확인]→ converted(판매전환)    ← 입금확인 알림톡 + offline_sales 자동생성·재고차감
  └─[취소]→ cancelled
```
전환 후(=offline_sales): 송장생성→발송 알림톡 / ALPS cron→배송완료 자동 / 약속✓→후기요청 자동. (EVENT 신규 로직 없음, 판매 인프라 재사용)

## 가격 (lib/event/pricing.ts = 서버 권위 / page_form.html JS 동일 복제)
- 품목 단가×수량 + 슬라이싱 가공(+20,000/자루).
- **묶음 할인**: 캠페인 `discount_rules[{unit_price,min_qty,bundle_price}]`. **같은 단가끼리**, min_qty 도달 시 **묶음 반복 + 나머지 정가**. 혼합 미적용.
  - 예 50000/3/130000 → 3자루=13만, 4자루=18만, 6자루=26만.
- 판매 전환 시 묶음할인 = `offline_sales.discount_amount`.

## 알림톡 (make-webhook.ts, webhook_consultation)
- 신규 3: `event_received`(접수확인·자동) / `event_payment_notice`(입금안내·사장님 버튼) / `event_payment_confirmed`(입금확인·자동).
- 발송/배송완료/후기 = 기존 판매 알림톡 재사용.
- ※ 솔라피 콘솔 3종 등록+검수 + Make 분기 필요. 미발송 시 [솔라피/Make/콘솔] 점검.

## 화면/코드
- 고객폼: `projects/event/page_form.html` (page.mamoru.kr, `?campaign=<id>`). 공개 API `app/api/event/public/{products,submit}`.
- 허브: `app/(dashboard)/events/page.tsx` (캠페인 카드 → 접수목록). hooks `use-events.ts`.
- 캠페인 API: `app/api/campaigns/route.ts`(GET/POST) + `[id]/route.ts`(PATCH). 접수 API `app/api/events/...`.
- 전환: `lib/event/convert-to-sale.ts`.
- 매장 워크인(사전접수 X) = `/sales/new`에서 EVENT 품목(category='EVENT') 선택 = 일반 판매. 특수처리 없음.

## 미래 (방향)
접수형(한정/공동구매/체험단/트레이드인)=캠페인 추가. 멤버십(서포터즈/앰버서더/테스터)=고객 등급(price_groups) 재사용, 별도 사이클.
