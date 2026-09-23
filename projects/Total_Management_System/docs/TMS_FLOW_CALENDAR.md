# 구글 캘린더 연동 흐름 (2026-09-24 정리)

> 사장님 캘린더(bsm@mamoru.kr `primary`)에 **확인이 필요한 것만** 올린다.
> 코드: `lib/google/` — `oauth.ts`(토큰) · `calendar-client.ts`(insert/patch/delete) · `event-formatter.ts`(형식 SSOT)
> · `calendar-sync.ts`(상담) · `repair-calendar-sync.ts`(복원수리 직접방문) · `shipping-todo-sync.ts`(송장 할 일)

## 1. 캘린더에 올라가는 것

| 종류 | 대상 | 시간 | 색 |
|---|---|---|---|
| 상담 매장방문 | confirmed 등 (취소 제외) | 방문시각 + 소요시간 | 초록 `2` |
| 상담 출장 | 〃 | 〃 | 보라 `3` |
| 고객 변경요청 | reschedule/change_requested | 〃 | 노랑 `5` |
| 복원수리 직접방문 | `proceed_type='직접방문'` | 방문시각 + 10분+자루당5분 | 주황 `6` (완료·취소 회색 `8`) |
| **송장 할 일** | 아래 2절 | **종일 · 한가함** | 회색 `8` |

## 2. 송장 할 일 (신규)

- **판매** `offline_sales`: `delivery_method='shipping'` AND `invoice_number IS NULL` AND 취소 아님
  - `pickup`(매장 수령)은 원래 송장이 없다 → 제외. 실측(09-24) shipping 66건 전부 송장 있음 = 오탐 0
- **납품** `deliveries`: `status='confirmed'` AND `tracking_number IS NULL` AND 취소 아님
  - `draft`(작성중)·`shipped`(직접전달 포함)·`cancelled` 제외
- **등록 다음 날**부터 올린다(당일 처리 흐름 방해 금지) · `calendar.shipping_todo_since`(기본 2026-09-24) 이후 등록분만
- 크론 `/api/cron/shipping-todo` **매시간**: 생성은 KST 08시만, **정리(삭제)는 매시간**
  - 송장 생성·출고·취소되면 최대 1시간 안에 일정이 사라진다
  - 판매/납품 API 마다 삭제 훅을 심지 않고 **여기서 일괄** — 경로 누락으로 유령 일정이 남는 구조를 피함
- 이벤트 id 보관: `system_settings.calendar.shipping_todo_events` = `{"sale:<id>":"<eventId>"}` (마이그레이션 없이 가동)
- 수동: `?dry=1` 대상만 · `?force=1` 시각 무시 생성 · `?asOf=YYYY-MM-DD` 기준일 대체

## 3. 형식 규칙 (event-formatter.ts = SSOT)

제목 `종류 · 이름 · 핵심하나` — 전화번호는 제목에서 뺀다(잘림). 확인 필요한 상태만 `[변경요청]` 접두.
```
출장 · 구교은 · 서초구          수리 · 김대현 · 1자루        납품 출고 · 나루토
```
본문은 4~5줄. 구분선(━)·항목 이모지·3줄 경고문·`tel:` 줄은 **전부 제거**(2026-09-24).
순서 = ① 연락처 ② 이 건의 내용(메모·수량) ③ 식별·접수일 ④ TMS 링크.
주소는 `location` 필드로만 (구글이 지도·길찾기로 띄움) — 본문에서 반복하지 않는다.
UUID 상담번호는 사람이 못 읽으므로 본문에서 생략. 출장 지역은 `address_sigungu`가 실측 전부 NULL이라 도로명에서 추출.

## 4. 형식을 바꿨을 때

이미 등록된 일정은 **상태가 바뀔 때만** 다시 써진다 → 앞으로 남은 일정을 한 번 훑는 경로:
`GET /api/cron/calendar-reformat` (크론 아님, 수동) · `?dry=1` 대상 · `?preview=1` 저장될 제목·본문 확인
