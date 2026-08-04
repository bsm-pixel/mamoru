# 직접방문(매장방문) 복원수리 알림톡 5종 — 솔라피 등록 + Make 모듈 완전 스펙

> 톤 = 기존 출장/상담과 동일: 마침표 없음 · ✅🔔 이모지 · • 중점 · `#{name}님, 안녕하세요` · 운영시간 푸터
> 변수 = TMS→Make payload **영문 키**와 일치 (한글명 아님)
> 코드 배선은 이미 완료(make-webhook 5종) — **솔라피 등록 + 카카오 검수 통과 후 실제 발송**

---

## Make로 넘어오는 payload 공통 구조 (필터·매핑 기준)

TMS `sendNotification` 이 Make 웹훅으로 보내는 JSON:
```json
{
  "_meta": { "func": "<EVENT>", "trigger": "tms" },
  "topic": "alrimtalk",
  "template": "<코드>",        // ← Router 필터에 이걸 사용 (예: "as_visit_booked")
  "event": "<EVENT>",          // ← 대문자 이벤트명도 함께 옴 (예: "AS_VISIT_BOOKED")
  "name": "고객명",
  "phone": "01012345678",
  "channel": "kakao",
  ... 각 변수 (as_id, visit_date, visit_time, qty, visit_duration_min, change_request_link)
}
```
- **Make Router 분기 필터**: `template` **equal to** `"<코드>"` (권장). `event` equal to `"<EVENT>"` 도 가능.
- **솔라피 템플릿 변수**: payload의 키 이름 그대로 `#{name}`, `#{visit_date}` … 로 매핑.

## 웹훅 시나리오 라우팅 (어느 Make 시나리오에 붙일지)

| 코드 | 웹훅 시나리오 | 이유 |
|---|---|---|
| `as_visit_booked` | **접수 알림** (webhook_as_received) | 접수건 → as_received와 같은 시나리오 |
| `as_visit_remind_24h` | **복원수리 상태변경** (webhook_repair) | 상태/후속 알림 |
| `as_visit_remind_2h` | 복원수리 상태변경 (webhook_repair) | 〃 |
| `as_visit_rescheduled` | 복원수리 상태변경 (webhook_repair) | 〃 |
| `as_visit_cancelled` | 복원수리 상태변경 (webhook_repair) | 〃 |

## 버튼값 공통 정의 (솔라피 버튼 등록 시)

| 버튼명 | 종류 | 값 (모바일·PC 공통) |
|---|---|---|
| 일정 확인·변경 | WL(웹링크) | `https://#{change_request_link}` (변수에 `page.mamoru.kr/projects/as/page_change_request.html?uid=<토큰>` 이 들어옴) |
| 매장 위치 보기 | WL(웹링크) | **정적** — 사장님 발급 네이버 지도 단축 URL |
| 1:1 문의 | BC(상담톡 전환) | MAMORU 카카오 채널 (출장/상담 템플릿의 "1:1 문의하기"와 동일) |
| 다시 예약하기 | WL(웹링크) | **정적** — 직접방문 접수 페이지 URL |

> ⚠️ 카카오 알림톡엔 "전화걸기" 버튼 타입이 없음 → 문의는 **BC(상담톡)** 로. (기존 상담 템플릿도 동일)
> 💡 **메시지 유형**: 접수완료·리마인드·변경 = **이미지형 추천**(헤더 = 매장 외관, 브랜드 일관성+위치 인지). 취소완료 = 텍스트형(절제). 이미지 규격 = 가로 800×400, 광고문구/전화번호 삽입 금지(검수 반려 사유).

### ✅ 버튼·주소 배치 규칙 (확정 2026-07-28)

| 템플릿 | 일정 확인·변경 | 매장 위치 보기 | 1:1 문의 | 다시 예약 | 본문 텍스트 주소 |
|---|:-:|:-:|:-:|:-:|:-:|
| ① 접수완료 | ✅ | ✅ | ✅ | | ✅ (첫 안내라 상세) |
| ② 리마인드 24h | ✅ | ✅ | | | ❌ |
| ③ 리마인드 2h | | ✅ **(피크)** | ✅ | | ❌ |
| ④ 변경안내 | ✅ | ✅ | | | ❌ |
| ⑤ 취소완료 | | | ✅ | ✅ | ❌ |

- **매장 위치 보기(지도)** = 방문하는 4종(①②③④) 모두. 특히 **2h 리마인드는 출발 직전 길찾기 피크라 필수**. 정적 링크라 비용 0.
- **본문 텍스트 주소** = 접수완료 ①에만(첫 안내). 리마인드·변경은 지도 버튼으로 대체(본문 간결). = 점진적 안내.

---

# ① as_visit_booked — 직접방문 접수완료

- **Make 필터**: `template` = `"as_visit_booked"` (event = `AS_VISIT_BOOKED`)
- **웹훅 시나리오**: 접수 알림 (webhook_as_received)
- **발송 트리거**: 직접방문 접수 즉시 (검수 후 submit 라우트 `as_received`→`as_visit_booked` 교체)
- **메시지 유형**: 이미지형 추천 (헤더=매장 외관)
- **변수**: `#{name}` `#{as_id}` `#{visit_date}` `#{visit_time}` `#{qty}` `#{visit_duration_min}` `#{change_request_link}`
- **버튼**: 일정 확인·변경(WL) · 매장 위치 보기(WL) · 1:1 문의(BC)

```
#{name}님, 안녕하세요
매장 방문 수리 접수가 완료되었어요

✅ 방문 예약 내역
• 접수번호 : #{as_id}
• 방문 일시 : #{visit_date} #{visit_time}
• 수리 수량 : #{qty}자루
• 예상 소요 : 약 #{visit_duration_min}분

🔔 방문 전 확인해 주세요
• 수리하실 가위를 꼭 지참해 주세요
• 일정 변경이나 취소는 아래 버튼에서 하실 수 있어요

매장 안내 : [매장주소]
운영시간 : 10:00 ~ 21:00
```

---

# ② as_visit_remind_24h — 방문 하루 전 (D-1)

- **Make 필터**: `template` = `"as_visit_remind_24h"` (event = `AS_VISIT_REMIND_24H`)
- **웹훅 시나리오**: 복원수리 상태변경 (webhook_repair)
- **발송 트리거**: 방문 전날 (신규 크론 `api/cron/repair-visit-remind`, 검수 후 배선)
- **메시지 유형**: 이미지형 추천 (헤더=매장 외관)
- **변수**: `#{name}` `#{visit_date}` `#{visit_time}` `#{qty}` `#{change_request_link}` (예상 소요는 접수완료에서 이미 안내 → 리마인드엔 생략)
- **버튼**: 일정 확인·변경(WL) · 매장 위치 보기(WL)

```
#{name}님, 안녕하세요
내일 매장 방문 수리 예약이 있어요

✅ 방문 예약 내역
• 방문 일시 : #{visit_date} #{visit_time}
• 수리 수량 : #{qty}자루

🔔 확인해 주세요
• 수리하실 가위를 잊지 말고 지참해 주세요
• 일정 변경이 필요하시면 아래 버튼을 눌러주세요

운영시간 : 10:00 ~ 21:00
```

---

# ③ as_visit_remind_2h — 방문 2시간 전 (당일)

- **Make 필터**: `template` = `"as_visit_remind_2h"` (event = `AS_VISIT_REMIND_2H`)
- **웹훅 시나리오**: 복원수리 상태변경 (webhook_repair)
- **발송 트리거**: 방문 당일 (크론 `api/cron/repair-visit-remind`, 검수 후 배선)
- **메시지 유형**: 이미지형 추천 (헤더=매장 외관)
- **변수**: `#{name}` `#{visit_time}` `#{qty}` `#{change_request_link}`
- **버튼**: 매장 위치 보기(WL) · 1:1 문의(BC)

```
#{name}님, 안녕하세요
오늘 매장 방문 시간이 곧 다가와요

✅ 오늘 방문 일정
• 방문 시간 : #{visit_time}
• 수리 수량 : #{qty}자루

🔔 안내 사항
• 수리하실 가위 지참을 다시 한번 확인해 주세요
• 매장 위치는 아래 버튼에서 확인하실 수 있어요

운영시간 : 10:00 ~ 21:00
```

---

# ④ as_visit_rescheduled — 예약 변경 안내 (고객 셀프 변경 시)

- **Make 필터**: `template` = `"as_visit_rescheduled"` (event = `AS_VISIT_RESCHEDULED`)
- **웹훅 시나리오**: 복원수리 상태변경 (webhook_repair)
- **발송 트리거**: 고객이 page_change_request 에서 일정 변경 완료 시 (API `resched` — **이미 배선됨**, 검수만 통과하면 발송)
- **메시지 유형**: 이미지형 추천 (헤더=매장 외관)
- **변수**: `#{name}` `#{as_id}` `#{visit_date}` `#{visit_time}` `#{qty}` `#{change_request_link}`
- **버튼**: 일정 확인·변경(WL) · 매장 위치 보기(WL)

```
#{name}님, 안녕하세요
매장 방문 수리 일정이 변경되었어요

✅ 변경된 방문 일정
• 방문 일시 : #{visit_date} #{visit_time}
• 수리 수량 : #{qty}자루

🔔 확인해 주세요
• 변경된 시간에 맞춰 방문해 주세요
• 추가 변경이나 취소는 아래 버튼에서 하실 수 있어요

운영시간 : 10:00 ~ 21:00
```

---

# ⑤ as_visit_cancelled — 예약 취소 완료 (고객 셀프 취소 시)

- **Make 필터**: `template` = `"as_visit_cancelled"` (event = `AS_VISIT_CANCELLED`)
- **웹훅 시나리오**: 복원수리 상태변경 (webhook_repair)
- **발송 트리거**: 고객이 page_change_request 에서 취소 확정 시 (API `cancel` — **이미 배선됨**, 검수만 통과하면 발송)
- **메시지 유형**: 텍스트형 권장 (취소엔 이미지 절제)
- **변수**: `#{name}` `#{as_id}` `#{visit_date}` `#{visit_time}`
- **버튼**: 다시 예약하기(WL) · 1:1 문의(BC)

```
#{name}님, 안녕하세요
매장 방문 수리 예약이 정상적으로 취소되었어요

✅ 취소된 예약 내역
• 방문 일정 : #{visit_date} #{visit_time}

· 다음에 더 좋은 상태로 뵙기를 바랄게요
· 수리가 다시 필요하시면 언제든 편하게 신청해 주세요

운영시간 : 10:00 ~ 21:00
```

---

## 검수 통과 후 TMS 배선 체크리스트

| # | 코드 | 발송 위치 | 상태 |
|---|---|---|---|
| ① | as_visit_booked | submit 라우트 직접방문 분기 → `as_visit_booked` + `change_request_link` (+ PUSH_CONFIG 접수푸시 유지) | **배선 완료 2026-08-04** |
| ② | as_visit_remind_24h | 신규 크론 `api/cron/repair-visit-remind` (10분간격) + 마이그 125 컬럼 | **배선 완료 2026-08-04** |
| ③ | as_visit_remind_2h | 동 크론 (0.5~2h 전) + 중복방지 컬럼 `visit_remind_2h_sent_at` | **배선 완료 2026-08-04** |
| ④ | as_visit_rescheduled | `api/repair/public/resched` (+ 리마인드 플래그 리셋·소요시간 공식 정렬) | **배선 완료** |
| ⑤ | as_visit_cancelled | `api/repair/public/cancel` | **배선 완료** |

> ✅ 2026-08-04: 전 5종 TMS 배선 완료. **마이그 125**(`visit_remind_24h_sent_at`/`visit_remind_2h_sent_at`) 실행 필요. 크론 `repair-visit-remind` `*/10 * * * *` 등록.

- Make 작업: **접수 알림 시나리오**에 `as_visit_booked` 분기 1개 / **복원수리 상태변경 시나리오**에 4종 분기(remind_24h/remind_2h/rescheduled/cancelled).
- 각 분기 = Router 필터(`template` equal to 코드) → 솔라피 모듈(해당 템플릿ID + 변수 매핑 + 버튼).

### ⚠️ 리마인드 크론 구축 시 필수 규칙 (옛 일정 리마인드 방지)
- 리마인드는 **레코드의 현재 `visit_date`/`visit_time` 기준**으로만 발송 → 변경 시 레코드가 덮이므로 옛 일정 리마인드는 원천 불가.
- **`resched`(일정 변경) 시 리마인드 발송 플래그(`visit_remind_24h_sent_at`·`visit_remind_2h_sent_at`)를 NULL로 리셋** → 변경된 새 일정에 리마인드가 다시 정상 발송. (크론 배선 시 resched 라우트에 추가)
- 크론 발송 조건: `proceed_type='직접방문'` + `status='intake'`(취소/수리시작 전) + 시간창 매칭 + 해당 플래그 NULL.
