# EVENT 알림톡 셋업 가이드 (솔라피 + Make) — 2026-06-26

> 이벤트(고객 접수) 자동 알림톡 3종. **코드는 이미 발송 호출 완성** — 솔라피 템플릿 등록·검수 + Make 분기만 추가하면 가동.
> 코드 전수 확인 기준(추측 X): `lib/notification/make-webhook.ts`, `api/event/public/submit/route.ts`, `api/events/[id]/route.ts`.

---

## 0. 전체 흐름 (사장님 버튼은 2번뿐)

| 단계 | 트리거 | 알림톡 | 발송 주체 |
|---|---|---|---|
| 1. 접수 | 고객이 폼 제출 | **① event_received** (접수확인) | 자동 |
| 2. 입금안내 | TMS EVENT허브 [입금안내] 클릭 (재고확인 후) | **② event_payment_notice** (계좌+금액) | 사장님 버튼 |
| 3. 입금확인 | TMS [입금확인 → 판매전환] 클릭 | **③ event_payment_confirmed** + 판매 자동생성 | 사장님 버튼 |
| 4. 발송/수령 | 판매관리 [택배발송] / [수령완료] | 기존 `sales_shipped` / 배송완료 자동 / 후기 자동 | 기존 인프라 (EVENT 신규 0) |

→ 4단계부터는 **판매로 전환된 뒤 기존 판매 알림톡**이 그대로 처리. EVENT 전용은 **①②③ 3종**만.

---

## 1. 알림톡 템플릿 3종 (솔라피 콘솔 등록용)

> 변수명은 **코드가 보내는 JSON 필드명과 100% 일치**해야 함(아래 표 그대로). 본문은 MAMORU 톤 초안 — 사장님 틀에 맞게 문구만 조정 가능, **변수명은 유지**.
> ⚠️ `#{total_amount}`는 콤마 없는 숫자("260000")로 들어옴. "260,000원" 원하면 → **Make에서 formatNumber 처리** 또는 코드 수정(별도).

### ① event_received — 접수 확인 (자동)
**보내는 변수**: `#{name}` `#{event_number}` `#{items}` `#{total_amount}` `#{receive_method}`
```
[MAMORU EVENT] 접수가 완료되었습니다

#{name}님, 신청해 주셔서 감사합니다.
아래 내용으로 접수되었습니다.

· 접수번호 : #{event_number}
· 신청품목 : #{items}
· 결제예정 : #{total_amount}원
· 수령방법 : #{receive_method}

재고 확인 후 입금 안내를 보내드립니다.
```
- 버튼: (선택) [채팅 상담] BC, extra `#{name}_event`

### ② event_payment_notice — 입금 안내 (사장님 [입금안내])
**보내는 변수**: `#{name}` `#{event_number}` `#{total_amount}`
> ⚠️ **계좌번호는 코드가 안 보냄 → 아래처럼 템플릿에 고정 텍스트로 입력**
```
[MAMORU EVENT] 입금 안내

#{name}님, 재고 확인이 완료되었습니다.
아래 계좌로 입금해 주시면 순차 처리됩니다.

· 접수번호 : #{event_number}
· 입금금액 : #{total_amount}원
· 입금계좌 : OO은행 000-000000-00000 (예금주 마모루)

입금자명이 신청자명과 다르면 채팅으로 알려주세요.
```
- 버튼: (선택) [채팅 상담] BC

### ③ event_payment_confirmed — 입금 확인 (사장님 [입금확인→판매전환])
**보내는 변수**: `#{name}` `#{event_number}` `#{total_amount}`
```
[MAMORU EVENT] 입금이 확인되었습니다

#{name}님, 입금이 정상 확인되었습니다. 감사합니다.

· 접수번호 : #{event_number}
· 결제금액 : #{total_amount}원

준비되는 대로 발송해 드리며,
발송 시 송장번호를 별도 알림톡으로 안내드립니다.
```
- 버튼: (선택) [채팅 상담] BC

---

## 2. Make 시나리오 — 어디에 / 필터 어떻게

### 어느 시나리오?
**기존 `consultation` 시나리오** (웹훅 = `MAKE_WEBHOOK_URL` = 설정 `notifications.webhook_consultation`).
→ 상담·후기·판매출고 알림톡이 쓰는 **그 웹훅 그대로**. 신규 웹훅 발급 **불필요**. 그 시나리오의 **라우터에 분기 3개만 추가**.

### 웹훅이 보내는 JSON (필터 기준 필드)
```json
{
  "topic": "alrimtalk",
  "template": "event_received",        ← ★ 필터는 이 필드
  "event": "EVENT_RECEIVED",           ← (대문자, 대체 가능)
  "name": "...", "phone": "0101234...",
  "event_number": "...", "items": "...", "total_amount": "260000", "receive_method": "택배발송"
}
```

### 라우터 분기 3개 (필터)
각 분기 **Filter** 조건 — 기존 분기들과 동일하게 **`1.template` Equal to** 사용:

| 분기 | Filter | 발송 솔라피 템플릿 | 매핑할 변수 |
|---|---|---|---|
| EVENT 접수 | `{{1.template}}` = `event_received` | event_received | name, event_number, items, total_amount, receive_method |
| EVENT 입금안내 | `{{1.template}}` = `event_payment_notice` | event_payment_notice | name, event_number, total_amount |
| EVENT 입금확인 | `{{1.template}}` = `event_payment_confirmed` | event_payment_confirmed | name, event_number, total_amount |

- 각 분기 = 기존 솔라피 발송 모듈 **복제** → 템플릿ID·변수 매핑만 교체.
- 변수 매핑: 솔라피 `#{name}` ← `{{1.name}}`, `#{event_number}` ← `{{1.event_number}}` … (이름 동일).
- (콤마 원하면) `#{total_amount}` 매핑값을 `{{formatNumber(1.total_amount; 0; ,; )}}` 식으로.

---

## 3. 가동 전 체크리스트

- [ ] **솔라피**: 위 3종 템플릿 등록 + 카카오 검수 신청 (검수 1~3영업일, 통과 전 발송 X)
- [ ] **Make**: consultation 시나리오 라우터에 분기 3개 추가(위 필터)
- [ ] **TMS 설정**: 알림톡 토글 ON — `notifications.event_received` / `event_payment_notice` / `event_payment_confirmed` (OFF면 코드가 발송 스킵)
- [ ] **계좌번호**: ② 템플릿 본문에 실제 계좌 고정 입력
- [ ] (이미 완료 가정) DB 마이그 103·104·105, EVENT 품목 등록, 캠페인 생성

## 4. 미발송 시 3단 점검 (알림톡 디버깅)
1. **TMS 설정 토글** ON 여부 + 코드 발송 로그(`[make-webhook] SKIP` 뜨면 웹훅 미설정)
2. **Make 시나리오** 분기 필터/변수 매핑 (template 철자)
3. **솔라피 템플릿** 검수 통과 여부 + 변수명 일치

관련: [reference_solapi_templates] · [project_event_system] · `docs/TMS_FLOW_EVENT.md`
