# 솔라피 알림톡 템플릿 — 아임웹 주문 프로세스

> 등록: 솔라피 콘솔 → 알림톡 템플릿 → 신규 등록
> 검수: 카카오 1~3 영업일
> 발송: Make 시나리오 `MAKE_WEBHOOK_URL` 분기

---

## 1. 주문 확인 (`order_confirmed`)

**트리거:** 결제 완료 후 (TMS sync 시 `pay_done` 최초 감지)

```
[마모루] #{name}님, 주문이 확인되었습니다.

#{name}님, 안녕하세요.
마모루를 이용해 주셔서 감사합니다.

■ 주문 내역
• 주문번호: #{order_no}
• 주문일시: #{order_date}
• 결제금액: #{paid_amount}원

■ 주문 상품
#{product_names}

배송 준비가 시작되면 다시 안내드리겠습니다.
```

| 버튼 | 타입 | URL |
|------|------|-----|
| 주문 조회 | WL | `https://mamoru.kr/mypage` |
| 1:1 문의 | WL | `https://mamoru.kr/counsel` |

**변수:**
- `name`: 주문자명
- `order_no`: 아임웹 주문번호
- `order_date`: 주문일시 (YYYY.MM.DD HH:mm)
- `paid_amount`: 결제금액 (콤마 포함)
- `product_names`: 상품명 (여러 개면 줄바꿈)

---

## 2. 발송 완료 (`order_shipped`)

**트리거:** 송장 생성 후 (TMS → 롯데 ALPS → 송장번호 확정)

```
[마모루] #{name}님, 주문하신 제품이 발송되었습니다.

#{name}님, 안녕하세요.
주문하신 제품의 배송이 시작되었습니다.

■ 배송 정보
• 택배사: #{courier}
• 송장번호: #{invoice_number}

■ 주문 상품
#{product_names}

배송 조회는 아래 버튼을 눌러 확인하실 수 있습니다.
```

| 버튼 | 타입 | URL |
|------|------|-----|
| 배송 조회 | WL | `https://www.lotteglogis.com/home/reservation/tracking/link498?InvNo=#{invoice_number}` |
| 1:1 문의 | WL | `https://mamoru.kr/counsel` |

**변수:**
- `name`: 주문자명
- `courier`: 택배사명 (롯데택배)
- `invoice_number`: 송장번호
- `product_names`: 상품명

---

## 3. 배송 완료 + 리뷰 요청 (`purchase_review_request`)

**트리거:** 배송완료 감지 (TMS sync 시 `delivered` 전환)

```
[마모루] #{name}님, 주문하신 제품은 만족스러우신가요?

#{name}님, 안녕하세요.
마모루 제품을 이용해 주셔서 감사합니다.

■ 주문 상품
#{product_names}

사용해 보신 후 간단한 후기를 남겨주시면
서비스 개선에 큰 도움이 됩니다.
```

| 버튼 | 타입 | URL |
|------|------|-----|
| 후기 남기기 | WL | `https://page.mamoru.kr/projects/reviews/page_review.html?uid=#{order_uid}&type=purchase&name=#{name}` |
| 1:1 문의 | WL | `https://mamoru.kr/counsel` |

**변수:**
- `name`: 주문자명
- `product_names`: 상품명 (여러 개면 콤마 구분)
- `order_uid`: TMS order.id (UUID)

---

## 4. 주문 취소 완료 (`order_cancelled`)

**트리거:** 취소 확정 (아임웹 → TMS sync 시 `cancelled` 전환)

```
[마모루] #{name}님, 주문 취소가 완료되었습니다.

#{name}님, 안녕하세요.
요청하신 주문 취소가 처리되었습니다.

■ 취소 내역
• 주문번호: #{order_no}
• 주문 상품: #{product_names}
• 환불금액: #{refund_amount}원

환불은 결제 수단에 따라 1~3 영업일 내 처리됩니다.
이용에 불편을 드려 죄송합니다.
```

| 버튼 | 타입 | URL |
|------|------|-----|
| 주문 조회 | WL | `https://mamoru.kr/mypage` |
| 1:1 문의 | WL | `https://mamoru.kr/counsel` |

**변수:**
- `name`: 주문자명
- `order_no`: 아임웹 주문번호
- `product_names`: 상품명
- `refund_amount`: 환불금액 (콤마 포함)

---

## 5. 교환/반품 접수 확인 (`order_refund_received`)

**트리거:** 환불 요청 감지 (TMS sync 시 `refund_request` 전환)

```
[마모루] #{name}님, 반품 요청이 접수되었습니다.

#{name}님, 안녕하세요.
반품 요청이 확인되어 안내드립니다.

■ 반품 내역
• 주문번호: #{order_no}
• 주문 상품: #{product_names}

반품 수거 및 환불 절차는 확인 후 별도 안내드리겠습니다.
궁금하신 사항은 1:1 문의를 이용해 주세요.
```

| 버튼 | 타입 | URL |
|------|------|-----|
| 주문 조회 | WL | `https://mamoru.kr/mypage` |
| 1:1 문의 | WL | `https://mamoru.kr/counsel` |

**변수:**
- `name`: 주문자명
- `order_no`: 아임웹 주문번호
- `product_names`: 상품명

---

## Make 시나리오 이벤트 매핑

| 템플릿 ID | Make 이벤트 | Webhook |
|----------|------------|---------|
| `order_confirmed` | `ORDER_CONFIRMED` | MAKE_WEBHOOK_URL |
| `order_shipped` | `ORDER_SHIPPED` | MAKE_WEBHOOK_URL |
| `purchase_review_request` | `PURCHASE_REVIEW_REQUEST` | MAKE_WEBHOOK_URL |
| `order_cancelled` | `ORDER_CANCELLED` | MAKE_WEBHOOK_URL |
| `order_refund_received` | `ORDER_REFUND_RECEIVED` | MAKE_WEBHOOK_URL |

---

## 1:1 문의 URL 확인 필요

현재 `https://mamoru.kr/counsel`로 기재했으나, 실제 상담 채널이 다르면 수정 필요:
- 카카오 채널 채팅: `https://pf.kakao.com/XXXX/chat`
- 아임웹 상담 페이지: 아임웹 URL
- GitHub Pages 상담 접수: `https://page.mamoru.kr/projects/consulting/page_counsel.html`

---

## 말투 규칙 (기존 템플릿과 통일)

- 헤더: `[마모루] #{name}님, ~` (한 문장 요약)
- 인사: `#{name}님, 안녕하세요.` (두 번째 줄)
- 섹션 구분: `■ 제목`
- 불릿: `•`
- 금액: `#{금액}원` (콤마 포함 숫자)
- 마무리: 다음 단계 안내 또는 감사 인사
- 버튼: 2개 (행동 버튼 + 1:1 문의)
- 이모지: 최소 사용 (본문에 넣지 않음)
