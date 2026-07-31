# EVENT 알림톡 셋업 가이드 (솔라피 + Make) — 2026-06-26

> 이벤트(고객 접수) 자동 알림톡. **2메시지 흐름 확정**(사장님 결정 2026-06-26): 접수완료(계좌 포함) → 입금확인.
> 코드 전수 확인 기준: `lib/notification/make-webhook.ts`, `api/event/public/submit/route.ts`, `api/events/[id]/route.ts`, `events/page.tsx`.

---

## 0. 전체 흐름 (2메시지 · 사장님 버튼 1번)

| 단계 | 트리거 | 알림톡 | 발송 |
|---|---|---|---|
| 1. 접수 | 고객 폼 제출 | **① event_received** (접수완료 + 계좌 + 입금요청) | 자동 |
| 2. 입금확인 | TMS EVENT허브 신규접수 → **[입금확인 → 판매전환]** | **② event_payment_confirmed** (입금완료) + 판매 자동생성·재고차감 | 사장님 버튼 |
| 3. 발송/수령 | 판매관리 [택배발송]/[수령완료] | 기존 `sales_shipped`/배송완료/후기 자동 | 기존 인프라 |

- **재고전환 이벤트 = 폼에서 재고 확정** → 접수 즉시 계좌 안내(입금요청) 가능. 별도 [입금안내] 단계 **생략**.
- **(선택)** 특정 건만 별도 입금안내가 필요하면 신규접수 화면 **[입금안내 별도 발송]** 보조버튼 → ③ event_payment_notice 발송(아래 3번 참고).
- ✅ 코드 보완 완료: 신규접수에서 바로 입금확인 가능(`events/page.tsx`), 접수완료에 `#{address}` 추가(`submit/route.ts`). **TMS 푸시 필요**(아직 미배포).

---

## 1. 알림톡 템플릿 (솔라피 콘솔 등록용 — 강조표기형)

> 변수명은 **코드가 보내는 JSON 필드명과 100% 일치**. `이벤트명`·`계좌`는 코드가 안 보내므로 **고정 텍스트**.
> ⚠️ `#{total_amount}`는 콤마 없는 숫자("130000"). 콤마 원하면 → Make에서 `{{formatNumber(...; 0; ,)}}`.
> 🆕 `#{items}`는 **품목마다 줄바꿈(\n)** 으로 들어감(2026-07-01, submit/route.ts). 템플릿에서 `// 주문 품목` 다음 줄에 `#{items}`를 두면 한 줄에 하나씩 표시됨. 묶음할인 때문에 품목별 가격은 넣지 않고 총액은 `#{total_amount}` 한 줄로.

### ① event_received — 접수완료 (자동) ★필수
**변수**: `#{name}` `#{items}` `#{total_amount}` `#{address}`
- 강조표기 제목: `주문 접수완료`
- 강조표기 보조문구: `MAMORU EVENT`
- 내용:
```
#{name}님 , 이벤트 품목 주문접수가
정상 처리되었습니다.

// 이벤트명 : 타사재고전환 이벤트
// 주문 품목
#{items}
// 입금금액 : #{total_amount}원
// 수령하실 주소 : #{address}

--------------------------------
🔔 하단 입금계좌에 입금주시면 됩니다.
입금확인 시 입금완료 안내 발송드리니 별도 입금확인 문의는 안주셔도 괜찮습니다 .
--------------------------------
우리)1002-439-462514 백성민
```
- 부가정보: `고객센터 운영시간: 오전 9시 ~ 오후 5시`

### ② event_payment_confirmed — 입금완료 (사장님 버튼) ★필수
**변수**: `#{name}` `#{event_number}` `#{total_amount}`
- 강조표기 제목: `입금 확인완료`
- 강조표기 보조문구: `MAMORU EVENT`
- 내용:
```
#{name}님 , 입금이 정상 확인되었습니다.
감사합니다.

// 이벤트명 : 타사재고전환 이벤트
// 접수번호 : #{event_number}
// 결제금액 : #{total_amount}원

--------------------------------
🔔 준비되는 대로 발송해 드리며,
발송 시 송장번호를 별도 알림톡으로 안내드립니다.
--------------------------------
```
- 부가정보: `고객센터 운영시간: 오전 9시 ~ 오후 5시`

### ③ event_payment_notice — 입금안내 (선택, 별도 발송용)
> 2메시지 흐름에선 **미사용**. 일부 건만 별도 계좌 안내가 필요할 때만 등록.
**변수**: `#{name}` `#{event_number}` `#{total_amount}`
- 강조표기 제목: `입금 안내` / 보조문구: `MAMORU EVENT`
- 내용: ①의 계좌 안내 부분과 동일 구성(접수번호·입금금액·계좌 고정).

---

## 2. Make 시나리오 — 어디에 / 필터

### 어느 시나리오?
**기존 `consultation` 시나리오**(웹훅 = `MAKE_WEBHOOK_URL` = 설정 `notifications.webhook_consultation`). 신규 웹훅 **불필요**. 라우터에 분기만 추가.

### 웹훅 JSON (필터 기준)
```json
{ "topic":"alrimtalk", "template":"event_received", "event":"EVENT_RECEIVED",
  "name":"...", "phone":"010...", "event_number":"...", "items":"...",
  "total_amount":"130000", "receive_method":"택배발송", "address":"..." }
```

### 라우터 분기 (필터 = `{{1.template}}` Equal to)
| 분기 | Filter | 솔라피 템플릿 | 매핑 변수 |
|---|---|---|---|
| 접수완료 ★ | `event_received` | event_received | name, items, total_amount, address |
| 입금완료 ★ | `event_payment_confirmed` | event_payment_confirmed | name, event_number, total_amount |
| (선택) 입금안내 | `event_payment_notice` | event_payment_notice | name, event_number, total_amount |

→ **2메시지 흐름 = ★ 2개만** 추가. 입금안내는 나중에 필요하면 추가.

---

## 3. 가동 전 체크리스트
- [ ] **솔라피**: ① event_received · ② event_payment_confirmed 등록 + 카카오 검수(1~3영업일). ②번 계좌·문구 확인
- [ ] **Make**: consultation 라우터에 분기 2개(★) 추가
- [ ] **TMS 설정**: 토글 ON — `notifications.event_received` / `event_payment_confirmed` (OFF면 발송 스킵)
- [ ] **TMS 배포(push)**: `submit/route.ts`(address), `events/page.tsx`(신규접수 입금확인 버튼) — **Vercel 빌드 필요**
- [ ] (이미) DB 마이그 103·104·105, EVENT 품목·캠페인

## 4. 미발송 3단 점검
1. TMS 설정 토글 ON + 발송 로그(`[make-webhook] SKIP`=웹훅 미설정)
2. Make 분기 필터/변수 매핑(template 철자)
3. 솔라피 검수 통과 + 변수명 일치

---

## 5. 발송(제품 발송) 알림톡 — **별도 EVENT 템플릿 불필요**

이벤트 흐름의 마지막 "제품 발송 알림톡"은 **새 템플릿을 안 만들어도 됩니다.**
- 입금확인 시 **판매 자동전환**(convertEventToSale → offline_sales) 되므로, 그 뒤부턴 **일반 판매 출고 흐름**을 탑니다.
- 판매관리에서 **택배발송(송장 생성) → 기사님 집하 → 자동 `sales_shipped`(출고 안내) 알림톡** (또는 수동 [출고완료]).
- 즉 **발송 알림톡 = 기존 `sales_shipped`** (이미 운영 중). EVENT 고객도 판매로 전환됐으니 동일하게 나갑니다.
- ⚠️ 조건: 판매전환된 건의 `customer_type`이 B2C여야 발송(B2B는 미발송). EVENT 고객은 B2C라 정상.

→ **EVENT용으로 새로 만들 알림톡은 접수완료 ① + 입금완료 ② 2개뿐** (+ 선택 ③ 입금안내). 발송은 판매 인프라 재사용.

---

## 6. Make 웹훅 만드는 법 (새 웹훅이 필요할 때만)

> ⚠️ **EVENT는 기존 consultation 웹훅을 쓰므로 새 웹훅 불필요**(2번 참고). 아래는 "새 시나리오를 따로 만들 때"용 참고.

1. 새 시나리오에서 **Webhooks → Custom webhook** 모듈 추가 → **Add** 클릭
2. **Create a webhook** 창:
   - **Webhook name**: 아무 이름 (예: `MAMORU EVENT`)
   - **API Key authentication**: **비워두기**(Add API key 안 함) — TMS는 `x-make-apikey` 헤더를 안 보내므로, 키를 넣으면 **전부 거부**됨. 반드시 빈 상태로.
   - **Save**
3. 생성되면 **웹훅 URL 복사**(`https://hook.eu2.make.com/…`)
4. 그 URL을 **TMS가 그 template을 쏘도록** 연결해야 함 → 이건 **코드 변경**(아래 7번)

---

## 7. 시나리오 결정 — 기존 consultation vs 신규 EVENT 전용

| | (A) 기존 consultation 시나리오에 분기 추가 | (B) 신규 EVENT 전용 시나리오 |
|---|---|---|
| TMS 코드 | **변경 없음** (event_* 이미 consultation 웹훅으로 감) | **필요** — `webhook_event` 설정 + make-webhook 라우팅 추가 |
| 작업량 | Make 분기 2개만 | 새 웹훅+시나리오 + TMS 배선·배포 |
| 깔끔함 | consultation에 이벤트 섞임 | EVENT만 분리(관리 명확) |
| **발송 알림톡** | consultation(sales_shipped) | **판매전환이라 어차피 consultation로 감** → EVENT 시나리오로 완전분리 안 됨 |

- **추천: (A) 기존 consultation 시나리오에 분기 2개 추가.** 지금 바로 되고(코드 0), 발송(sales_shipped)도 어차피 consultation이라 분리 실익이 적음.
- **(B)를 원하면**: 사장님이 "EVENT 시나리오 분리로 가자" 주시면 → 내가 `webhook_event` 설정 + 라우팅 배선(접수/입금만 분리, 발송은 판매라 유지). 새 웹훅은 위 6번대로 생성.

---

## 8. 지금 사장님이 하실 일 (A안 기준)
1. **솔라피 등록 + 검수**: ① event_received · ② event_payment_confirmed (양식은 1번 참고)
2. **Make**: **consultation 시나리오** Router에 분기 2개 추가
   - `{{template}}` = `event_received` → 솔라피 event_received (변수: name·items·total_amount·address)
   - `{{template}}` = `event_payment_confirmed` → 솔라피 event_payment_confirmed (변수: name·event_number·total_amount)
3. **TMS 설정 토글 ON**: `notifications.event_received` · `event_payment_confirmed`
4. **테스트**: 나에게 "이벤트 접수 테스트 쏴줘" 하면 각 template payload를 UTF-8로 발송해 드림
5. 발송 알림톡은 판매전환 후 자동(sales_shipped) — 별도 작업 없음

관련: [reference_solapi_templates] · [project_event_system] · `docs/TMS_FLOW_EVENT.md` · `projects/Total_Management_System/docs/MANUAL_EVENT.md`
