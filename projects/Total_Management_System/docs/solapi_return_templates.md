# 솔라피 반품·교환수거 알림톡 템플릿 (2026-08-25)

> TMS 코드는 이미 발송 준비됨(`make-webhook.ts` return_received / return_inbound).
> 아래 2개를 **솔라피 콘솔에 등록(검수 1~3영업일)** + **Make에 분기 추가**하면 고객 알림톡이 실제 발송됩니다.
> 등록 전에도 **사장님 폰 푸시는 이미 정상 발송**됩니다.
>
> 변수는 TMS가 넘기는 `data` 키와 1:1: `#{name}` `#{return_number}` `#{product_name}` `#{pickup_method}` `#{pickup_date}`
> (Make 시나리오에서 웹훅 payload의 name/data.* 를 이 변수에 매핑)

---

## 1) `return_received` — 반품·교환수거 접수 (수거 접수 직후 고객에게)

**템플릿명(코드)**: `return_received`
**Make 이벤트명**: `RETURN_RECEIVED` (make-webhook TEMPLATE_EVENT_MAP과 일치)

```
[마모루] #{name}님, 교환/반품을 위한 제품 수거를 접수했습니다.

■ 접수 내역
• 접수번호: #{return_number}
• 제품: #{product_name}
• 회수 방식: #{pickup_method}

■ 안내
회수 방식에 맞춰 제품을 준비해 주세요.
제품이 마모루에 도착하면 다시 안내드리겠습니다.

문의사항은 아래 버튼으로 편히 연락 주세요.
```

**버튼**
- `1:1 문의하기` (WL) → 카카오 채널 상담 링크

> 강조 표기(BC/EX) 필요 시 상단에 `#{name}님 수거 접수` 정도로. 50바이트(한글 약 16자) 이내.

---

## 2) `return_inbound` — 반품 입고완료 (제품 도착·입고 처리 시 고객에게)

**템플릿명(코드)**: `return_inbound`
**Make 이벤트명**: `RETURN_INBOUND`

```
[마모루] #{name}님, 보내주신 제품이 안전하게 도착했습니다.

■ 확인 내역
• 접수번호: #{return_number}
• 제품: #{product_name}

확인 후 교환/반품 처리를 진행해 드리겠습니다.
완료되면 다시 안내드리겠습니다.

감사합니다.
```

**버튼**
- `1:1 문의하기` (WL) → 카카오 채널 상담 링크

---

## 등록 후 확인
- 솔라피 검수 통과 → Make 시나리오에 `RETURN_RECEIVED` / `RETURN_INBOUND` 분기 추가(기존 as_* 분기 복제) → 웹훅 URL을 make-webhook이 쓰는 consultation 웹훅으로.
- 알림톡 on/off는 TMS 설정의 `notifications.return_received` / `notifications.return_inbound` (키 없으면 기본 ON).
- 테스트: `/returns`에서 반품 접수/입고완료 → 고객 폰 문자 수신 확인.
