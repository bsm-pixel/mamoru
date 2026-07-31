# EVENT 알림톡 셋업 가이드 (솔라피 + Make) — EVENT 전용 시나리오

> 이벤트(고객 접수) 자동 알림톡. **EVENT 전용 시나리오로 분리**(사장님 결정 2026-07-31):
> 접수확인(비용안내) → 입금확인 → **출고완료(EVENT 전용)**. 상담 시나리오와 완전 분리 = 모듈 헷갈림 방지.
> 코드 전수 확인 기준: `lib/notification/make-webhook.ts`, `lib/notification/sales-shipped.ts`, `lib/event/convert-to-sale.ts`, `api/event/public/submit/route.ts`, `api/events/[id]/route.ts`.

---

## 0. 전체 흐름 (3메시지 · 사장님 버튼 1번)

| 단계 | 트리거 | 알림톡 | 발송 |
|---|---|---|---|
| 1. 접수 | 고객 폼 제출 | **① event_received** (접수확인 + 비용안내 + 계좌) | 자동 |
| 2. 입금확인 | TMS EVENT허브 신규접수 → **[입금확인 → 판매전환]** | **② event_payment_confirmed** (입금완료) + 판매 자동생성·재고차감 | 사장님 버튼 |
| 3. 출고완료 | 판매관리 [택배발송] → 기사님 집하 감지(자동) 또는 [출고완료](수동) | **③ event_shipped** (EVENT 전용 출고완료 · 송장) | 자동 |

- **③ 출고완료는 이제 EVENT 전용 `event_shipped`** — 일반 판매의 `sales_shipped` 와 **다른 필터·다른 경로**.
- EVENT 접수페이지로 들어온 주문(=판매전환분)만 자동으로 `event_shipped` 로 나가고, 일반 판매는 그대로 `sales_shipped`.
  - 판별 기준: 판매 memo 가 `EVENT 전환 …`(= `convertEventToSale` 이 남기는 값)으로 시작하는 건 → `event_shipped`.
  - LS(재고판매)는 `재고판매 전환`, 수동 판매는 사용자 입력이라 **오발송 없음**.

---

## 1. 알림톡 템플릿 (솔라피 콘솔 등록용 — 강조표기형)

> **톤 = MAMORU "조용히 압도하는 전문가"** — 절제·정중·장인의 결. 과장·이모지·군더더기 배제, 존댓말 유지.
> 변수명은 **코드가 보내는 JSON 필드명과 100% 일치**. `계좌`는 코드가 안 보내므로 **고정 텍스트**.
> ⚠️ `#{total_amount}`는 콤마 없는 숫자("130000"). 콤마 원하면 → Make에서 `{{formatNumber(...; 0; ,)}}`.
> 🆕 `#{items}`는 **품목마다 줄바꿈(\n)**. 템플릿에서 `주문 내역` 다음 줄에 `#{items}`.

### ① event_received — 접수확인 + 비용안내 (자동) ★필수
**변수**: `#{name}` `#{items}` `#{total_amount}` `#{address}`
- 강조표기 제목: `주문이 접수되었습니다` / 보조문구: `MAMORU`
- 내용:
```
#{name}님, 주문을 정확히 확인했습니다.

주문 내역
#{items}

결제 금액 : #{total_amount}원
받으실 곳 : #{address}

아래 계좌로 입금해 주시면
확인 후 바로 준비를 시작하겠습니다.
입금 확인은 별도 안내로 보내드리니
따로 문의 주지 않으셔도 괜찮습니다.

우리은행 1002-439-462514  백성민
```
- 부가정보: `고객센터 · 평일 09:00 ~ 17:00`

### ② event_payment_confirmed — 입금확인 (사장님 버튼) ★필수
**변수**: `#{name}` `#{event_number}` `#{total_amount}`
- 강조표기 제목: `입금이 확인되었습니다` / 보조문구: `MAMORU`
- 내용:
```
#{name}님, 입금을 확인했습니다.
믿고 맡겨 주셔서 감사합니다.

접수번호 : #{event_number}
결제 금액 : #{total_amount}원

지금부터 정성껏 준비하여
가장 좋은 상태로 보내드리겠습니다.
발송이 완료되면 송장번호를
다시 안내해 드리겠습니다.
```
- 부가정보: `고객센터 · 평일 09:00 ~ 17:00`

### ③ event_shipped — 출고완료 (자동, EVENT 전용) ★필수 · 신규
**변수**: `#{name}` `#{id}` `#{goods_name}` `#{tracking}` `#{courier}`
- 강조표기 제목: `상품이 발송되었습니다` / 보조문구: `MAMORU`
- 내용:
```
#{name}님, 주문하신 상품을
방금 발송했습니다.

주문번호 : #{id}
상품 : #{goods_name}
택배사 : #{courier}
송장번호 : #{tracking}

송장번호로 배송 조회가 가능합니다.
받아보신 뒤 짧은 사용 후기를 남겨 주시면
다음을 만드는 데 큰 힘이 됩니다.
```
- 부가정보: `고객센터 · 평일 09:00 ~ 17:00`

> ⚠️ `#{id}` = 판매번호(OS-…) 또는 판매 UUID(폴백). `#{tracking}` = 롯데 송장번호. `#{courier}` 기본 "롯데택배".

### (선택) event_payment_notice — 입금안내 별도 발송
> 3메시지 흐름에선 미사용. 특정 건만 계좌를 다시 안내할 때만 등록. 변수 `#{name}` `#{event_number}` `#{total_amount}`.
- 강조표기 제목: `입금 안내` / 보조문구: `MAMORU`
- 내용: ①의 계좌 안내 부분과 동일 구성(접수번호·결제 금액·계좌 고정 텍스트).

---

## 2. Make 시나리오 — 신규 EVENT 전용

### 어느 시나리오?
**신규 `MAMORU EVENT` 시나리오**(웹훅 = 설정 `notifications.webhook_event` = env `MAKE_EVENT_WEBHOOK_URL`).
- 상담 시나리오에 있던 **기존 EVENT 모듈(접수완료·입금확인)은 삭제**한다. 출고완료는 원래 `sales_shipped`(상담 시나리오)로 나가던 걸 → 이제 `event_shipped`(EVENT 시나리오)로 분리.
- 웹훅 미설정(빈칸)이면 코드가 **상담 웹훅으로 폴백**하므로, 전환 순서는 아래 6번을 지킬 것(웹훅 세팅 → 그다음 상담 모듈 삭제).

### 웹훅 JSON (필터 기준)
```json
{ "topic":"alrimtalk", "template":"event_shipped", "event":"EVENT_SHIPPED",
  "name":"...", "phone":"010...", "id":"OS-...", "goods_name":"블런트 5.5인치",
  "tracking":"1234567890", "courier":"롯데택배" }
```

### 라우터 분기 (필터 = `{{1.template}}` Equal to)
| 분기 | Filter (`template`) | 솔라피 템플릿 | 매핑 변수 |
|---|---|---|---|
| 접수확인 ★ | `event_received` | event_received | name, items, total_amount, address |
| 입금확인 ★ | `event_payment_confirmed` | event_payment_confirmed | name, event_number, total_amount |
| 출고완료 ★ | `event_shipped` | event_shipped | name, id, goods_name, tracking, courier |
| (선택) 입금안내 | `event_payment_notice` | event_payment_notice | name, event_number, total_amount |

---

## 3. Make 웹훅 만드는 법 (신규)

1. 새 시나리오에서 **Webhooks → Custom webhook** 모듈 추가 → **Add**
2. **Create a webhook** 창:
   - **Webhook name**: `MAMORU EVENT`
   - **API Key authentication**: **비워두기** — TMS는 `x-make-apikey` 헤더를 안 보내므로, 키를 넣으면 **전부 거부**됨. 반드시 빈 상태.
   - **Save**
3. 생성된 **웹훅 URL 복사**(`https://hook.eu2.make.com/…`)
4. TMS **설정 → 알림 → Make 웹훅 URL (이벤트)** 칸에 붙여넣기 + 저장 (또는 Vercel env `MAKE_EVENT_WEBHOOK_URL`)

---

## 4. 가동 전 체크리스트
- [ ] **솔라피**: ① event_received · ② event_payment_confirmed · ③ **event_shipped(신규)** 등록 + 카카오 검수(1~3영업일)
- [ ] **Make**: 신규 `MAMORU EVENT` 시나리오 + 웹훅 생성 + Router 분기 3개(★)
- [ ] **TMS 설정**: **Make 웹훅 URL (이벤트)** 에 신규 웹훅 붙여넣기
- [ ] **(전환 순서)** 웹훅 세팅 확인 후 → 상담 시나리오의 **기존 EVENT 모듈 삭제**
- [ ] 토글: `notifications.event_received` / `event_payment_confirmed` / `event_shipped` — **키 없으면 발송(fail-open)**. 끄고 싶을 때만 false
- [ ] (이미) DB 마이그 103·104·105, EVENT 품목·캠페인

---

## 5. 미발송 3단 점검
1. TMS 설정 — **이벤트 웹훅 URL** 채워졌는지 (`[make-webhook] SKIP … webhook_event 미설정`=빈칸+상담도 빈칸)
2. Make 분기 필터/변수 매핑(template 철자: `event_received`/`event_payment_confirmed`/`event_shipped`)
3. 솔라피 검수 통과 + 변수명 일치

---

## 6. 전환 순서 (기존 → EVENT 전용, 유실 없이)
1. 솔라피 event_shipped 등록·검수 통과
2. Make 신규 `MAMORU EVENT` 시나리오 + 웹훅 생성 + 분기 3개
3. TMS 설정에 **이벤트 웹훅 URL** 저장 → 이 시점부터 event_* 는 EVENT 시나리오로 감
4. 상담 시나리오에서 **기존 EVENT 모듈(접수완료·입금확인) 삭제** — 출고완료(sales_shipped)는 일반 판매용이라 **그대로 둠**
5. 테스트: 나에게 "이벤트 접수/입금/출고 테스트 쏴줘" → 각 template 을 UTF-8로 발송

> ⚠️ 3번을 건너뛰고 4번(모듈 삭제)부터 하면, 웹훅 빈칸 폴백으로 상담 시나리오에 갔다가 모듈이 없어 유실됨. **반드시 3 → 4 순서.**

---

## 7. 코드 배선 요약 (2026-07-31)
- `make-webhook.ts` — `webhook_event`(설정/env `MAKE_EVENT_WEBHOOK_URL`, 미설정 시 consultation 폴백) + `EVENT_TEMPLATES` 라우팅 + `event_shipped` 템플릿/EVENT맵/토글키
- `sales-shipped.ts` — 출고 발송 직전 판매 memo `EVENT 전환` 접두어면 `event_shipped`, 아니면 `sales_shipped`. 집하 자동(cron)·수동 [출고완료] **두 경로 공통**(단일 함수)
- `notification-settings.tsx` — 설정에 **Make 웹훅 URL (이벤트)** 입력칸 추가

관련: [reference_solapi_templates] · [project_event_system] · `docs/TMS_FLOW_EVENT.md` · `projects/Total_Management_System/docs/MANUAL_EVENT.md`
