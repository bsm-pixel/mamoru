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

---

## 3) `adjust_shipped` — 점검·조정 후 재발송 (2026-11-12) · 템플릿 이름 `TMS-점검조정발송`

- **언제**: 판매 제품을 회수해 점검·조정한 뒤 **같은 제품**을 다시 보낼 때, 재발송 송장이 **기사님 집하 스캔**되면 1회 자동(크론 [3-C]). 접수·입고 땐 안 보냄.
- **Make**: 반품·교환 시나리오(webhook_returns)에 분기 1개 — 조건 `template` = `adjust_shipped` (event `ADJUST_SHIPPED`). 기존 `exchange_shipped` 분기 복제 후 템플릿ID만 바꿈. 새 시나리오면 `Store incomplete executions` = Yes.
- **카테고리**: 배송 > 상태안내 (006001) · **이미지형**
- **배너**: **점검·조정 완료 / 다시 보내드려요** + 오른쪽 3D 아이콘(상자 위 가위 + 체크)
- 🚫 「수리」「주문」 단어 금지(새 제품 산 고객이 "수리했다고?" 오해) · 리뷰 요청 없음

**본문**
```
#{name}님, 안녕하세요
말씀해 주신 부분 꼼꼼히 점검하고
조정을 마쳐 다시 보내 드렸습니다

📦 발송 내역
• 제품 : #{product_name}
• 택배사 : #{courier}
• 송장번호 : #{tracking}

받아 보시고 불편한 점이 있으면
언제든 편하게 말씀해 주세요
```

**버튼**: 버튼 종류 **「배송조회」** — 링크 없음 (카카오 자동 배송조회)

**변수**(TMS가 보내는 키 그대로): `name`(고객명) · `product_name`(제품명) · `courier`(택배사, 보통 롯데택배) · `tracking`(송장번호) · `return_number`(접수번호 — 본문 미사용, 필요 시)

**검수 요청 메모**: "판매한 제품을 고객 요청으로 회수해 점검·조정한 뒤 재발송할 때, 택배 집하 시점에 1회 발송되는 배송 상태 안내입니다. 광고 내용 없음."
