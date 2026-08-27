# 솔라피 반품·교환수거 알림톡 템플릿 (2026-08-25)

> TMS 코드는 이미 발송 준비됨(`make-webhook.ts` return_received / return_inbound).
> 아래 2개를 **솔라피 콘솔에 등록(검수 1~3영업일)** + **Make에 분기 추가**하면 고객 알림톡이 실제 발송됩니다.
> 등록 전에도 **사장님 폰 푸시는 이미 정상 발송**됩니다.
>
> 기존 주문 알림톡(`order_confirmed` 등)과 동일 톤. 변수는 TMS가 넘기는 `data` 키와 1:1:
> `#{name}` `#{return_number}` `#{product_name}` `#{pickup_method}`
> (Make 시나리오에서 웹훅 payload의 name/data.* 를 이 변수에 매핑)
> ※ 이모지는 헤드라인 1개만(카카오 검수 안전). 본문 과다 이모지는 반려 위험.

---

## 1) `return_received` — 반품·교환수거 접수 (수거 접수 직후)

**Make 이벤트명**: `RETURN_RECEIVED`

```
[마모루] #{name}님, 반품·교환 수거가 접수되었습니다. 📦

#{name}님, 안녕하세요.
반품·교환을 위한 제품 수거가 접수되었습니다.

■ 접수 내역
• 접수번호: #{return_number}
• 제품: #{product_name}
• 회수 방식: #{pickup_method}

회수 방식에 맞춰 제품을 준비해 주세요.
제품이 도착하면 다시 안내드리겠습니다.
```

| 버튼 | 타입 | URL |
|------|------|-----|
| 1:1 문의 | WL | `https://mamoru.kr/counsel` |

**변수**: `name`(고객명) · `return_number`(접수번호) · `product_name`(제품명) · `pickup_method`(방문수거/택배수거/직접반납)

---

## 2) `return_inbound` — 반품 입고완료 (제품 도착·입고 처리 시)

**Make 이벤트명**: `RETURN_INBOUND`

```
[마모루] #{name}님, 보내주신 제품이 도착했습니다. ✅

#{name}님, 안녕하세요.
반품·교환 제품이 마모루에 안전하게 도착했습니다.

■ 확인 내역
• 접수번호: #{return_number}
• 제품: #{product_name}

확인 후 처리를 진행해 드리며,
완료되면 다시 안내드리겠습니다.

감사합니다.
```

| 버튼 | 타입 | URL |
|------|------|-----|
| 1:1 문의 | WL | `https://mamoru.kr/counsel` |

**변수**: `name`(고객명) · `return_number`(접수번호) · `product_name`(제품명)

---

## 등록 후 확인
- 솔라피 검수 통과 → Make 시나리오에 `RETURN_RECEIVED` / `RETURN_INBOUND` 분기 추가(기존 as_* 분기 복제) → 웹훅 URL은 make-webhook이 쓰는 consultation 웹훅.
- 알림톡 on/off는 TMS 설정 `notifications.return_received` / `notifications.return_inbound` (키 없으면 기본 ON).
- 테스트: `/returns`에서 접수/입고완료 → 고객 폰 문자 수신 확인. (디버깅은 [reference_solapi_templates] 3단계: 코드→Make→솔라피)
