# TMS 흐름도 — EVENT (고객 접수 이벤트)

> 최종 업데이트: 2026-09-13 · 알림톡 범용화(이벤트명·고객 안내 문구 변수) / (최초) 2026-06-15 재고 전환 이벤트 1탄으로 구축. 접수형 이벤트 공통 허브.

## 개요
릴스/DM → **EVENT 접수 페이지(카탈로그형)** → 접수 → 입금안내 → 입금확인 → **판매(offline_sales) 자동 전환** → 발송/배송완료/후기(기존 판매 인프라). 캠페인(이벤트)별로 분리 관리.

## 데이터 모델
- `event_campaigns` (마이그 104·105): 이벤트 단위. `type`(stock_clearance/limited/group_buy/tester/trade_in/other), `status`(active/ended), `is_default`, `discount_rules` jsonb, `customer_notice`(마이그 145 — 신청완료 알림톡 안내 문구), `payment_type`·`items_label`(마이그 152 — 무료 이벤트·신청항목 표기).
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

### 2026-09-16 무료(증정·체험단) 이벤트 지원 — 마이그 `152`

**문제**: 범용 템플릿이 "유료 주문"을 전제로 굳어 있었다.
- `EVENT_신청완료` 본문에 `결제 금액`·`받으실 곳`·`입금 계좌`가 **고정 텍스트** → 증정인데 **"결제 금액 0원"**이 떠서 고객이 "왜 결제가?" 하게 된다
- `#{items}` 는 항상 "제품명 N개" → 체험단인데 제품명이 노출
- 🚨 더 근본: **판매 전환이 `confirm_payment` 한 곳뿐** → 무료 이벤트도 [입금확인]을 눌러야 진행되고, 그 순간 **"입금이 확인되었습니다 / 0원"** 알림톡이 나갔다

**해결**: 캠페인이 성격을 들고 있고, 고객에게 보이는 문구는 **TMS가 조립**한다.

> 🔑 알림톡은 **변수 안의 변수를 치환하지 않는다**(치환 1회). `#{event_notice}` 값에 `#{address}` 를 넣어 보내면 **글자 그대로** 나간다.
> 그래서 주소·금액·계좌를 끼워 넣으려면 **서버가 문자열을 완성해서** 한 변수로 넘겨야 한다 → 그게 `event_detail`.

| 캠페인 컬럼 (152) | 뜻 |
|---|---|
| `payment_type` | `paid`(입금 필요) / `free`(증정·체험단) |
| `items_label` | `#{items}` 대체 문구(예: 체험단 신청). 비우면 제품명+수량. **내부 품목·재고·판매전환엔 영향 없음** |

**`#{event_detail}` 조립 규칙** (`lib/event/campaign-notify.ts`)

| 상황 | 내용 |
|---|---|
| 유료(금액>0) | 결제 금액 + 수령지 + **입금 계좌** + "입금이 확인되면 바로 준비를 시작합니다" |
| 무료 | 수령지 + 캠페인 안내문구(없으면 "바로 준비를 시작합니다") — **금액·계좌 줄이 아예 생기지 않음** |
| 유료인데 금액 0 | 무료와 같게 처리 (할인으로 0원이 된 경우) |
| 매장 수령 | `• 수령 방법 : 매장 방문 수령` |

- **계좌가 변수 안으로 들어왔다** — 전엔 알림톡 본문 고정 텍스트라 계좌 변경 시 **본문 수정 + 재검수(1~3영업일)** 가 필요했다. 이제 `EVENT_BANK_ACCOUNTS` 상수만 고치면 되고 무료 이벤트엔 아예 안 나간다
- **무료 캠페인은 `event_payment_confirmed` 알림톡을 보내지 않는다.** 버튼도 [입금확인 → 판매 전환] → **[신청 확정 → 발송 준비]** 로 바뀐다. 판매 전환·재고 차감은 그대로(금액 0) — 무료 이벤트의 첫 자동 알림은 **출고완료(`event_shipped`)**
- ⚠️ 본문을 통째로 변수로 만들면 카카오 심사에서 "치환문구로만 구성"으로 반려될 수 있어, 인사·신청내역·안내 **골격은 고정 텍스트로 남겼다**

## 화면/코드
- 고객폼: `projects/event/page_form.html` (page.mamoru.kr, `?campaign=<id>`). 공개 API `app/api/event/public/{products,submit}`.
- 허브: `app/(dashboard)/events/page.tsx` (캠페인 카드 → 접수목록). hooks `use-events.ts`.
- 캠페인 API: `app/api/campaigns/route.ts`(GET/POST) + `[id]/route.ts`(PATCH). 접수 API `app/api/events/...`.
- 전환: `lib/event/convert-to-sale.ts`.
- 매장 워크인(사전접수 X) = `/sales/new`에서 EVENT 품목(category='EVENT') 선택 = 일반 판매. 특수처리 없음.

## 미래 (방향)
접수형(한정/공동구매/체험단/트레이드인)=캠페인 추가. 멤버십(서포터즈/앰버서더/테스터)=고객 등급(price_groups) 재사용, 별도 사이클.
