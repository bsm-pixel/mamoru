# 직접방문(매장방문) 복원수리 알림톡 템플릿 5종 — 솔라피 등록용

> 2026-07-28 작성 · 07-28 재작성(톤·변수 정정).
> 톤 = **기존 출장/상담 템플릿과 동일**: 마침표 없음 · ✅🔔 이모지 · • 중점 · `#{name}님, 안녕하세요` 시작 · 운영시간 푸터.
> 변수 = **TMS→Make payload 영문 키와 일치**(한글명 아님). Make 매핑 그대로 흐름.
> 검수 통과 후 TMS 코드 배선(make-webhook + 리마인드 크론 + 변경/취소 API).

## 변수 키 (payload 실측 — submit/route.ts + 변경/취소 API)

| 템플릿 변수 | 값(TMS가 채움) |
|---|---|
| `#{name}` | 고객명 |
| `#{as_id}` | 접수번호 |
| `#{visit_date}` | 방문일 (한글 포맷, 예: 7월 30일) |
| `#{visit_time}` | 방문시간 (HH:MM) |
| `#{qty}` | 수리 수량(마모루+타사) |
| `#{visit_duration_min}` | 예상 소요(분, 30/60) |
| `#{change_request_link}` | 일정확인/변경 페이지 URL (uid 포함) |

- **매장 주소**는 payload 변수가 아님 → 본문에 **정적으로 직접 입력**(아래 `[매장주소]` 자리) + "매장 위치 보기" 버튼(정적 네이버 지도 URL).
- 버튼값: 일정확인/변경(WL, `#{change_request_link}`) · 매장 위치 보기(WL 정적) · 전화 문의(DS 대표번호) · 다시 예약하기(WL 접수페이지)

---

## ① as_visit_booked — 직접방문 접수완료 (접수 즉시)
변수: `name` `as_id` `visit_date` `visit_time` `qty` `visit_duration_min` `change_request_link`
버튼: 일정확인/변경 · 매장 위치 보기 · 전화 문의

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

## ② as_visit_remind_24h — 방문 하루 전 (D-1)
변수: `name` `visit_date` `visit_time` `qty` `visit_duration_min` `change_request_link`
버튼: 일정확인/변경 · 매장 위치 보기 · 전화 문의

```
#{name}님, 안녕하세요
내일 매장 방문 수리 예약이 있어요

✅ 방문 예약 내역
• 방문 일시 : #{visit_date} #{visit_time}
• 수리 수량 : #{qty}자루
• 예상 소요 : 약 #{visit_duration_min}분

🔔 확인해 주세요
• 수리하실 가위를 잊지 말고 지참해 주세요
• 일정 변경이 필요하시면 아래 버튼을 눌러주세요

운영시간 : 10:00 ~ 21:00
```

---

## ③ as_visit_remind_2h — 방문 2시간 전 (당일)
변수: `name` `visit_time` `qty` `change_request_link`
버튼: 매장 위치 보기 · 전화 문의

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

## ④ as_visit_rescheduled — 예약 변경 안내 (고객 셀프 변경 시)
변수: `name` `visit_date` `visit_time` `qty` `change_request_link`
버튼: 일정확인/변경 · 매장 위치 보기 · 전화 문의

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

## ⑤ as_visit_cancelled — 예약 취소 완료 (고객 셀프 취소 시)
변수: `name` `visit_date` `visit_time`
버튼: 다시 예약하기 · 전화 문의

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

## 일정확인/변경 페이지 & 흐름 (상담 page_change_request.html 미러링)

- **링크**: `https://page.mamoru.kr/projects/as/page_change_request.html?uid=#{as_uid}`
- **페이지**: 예약 정보 카드 + [일정 변경] [예약 취소]
  - **[일정 변경]** → 방문일/시간 즉시 변경 모달 → 변경 접수 → `as_visit_rescheduled` 발송 → 카톡 인앱 창 닫힘
  - **[예약 취소]** → 재확인 모달(일정 유지하기 / 취소확정하기) → 취소확정 → `status='cancelled'` + `as_visit_cancelled` 발송 → 완료 화면 → 창 닫힘
- **카카오 인앱**: 나가기·완료 후 [reference_kakao_inapp_close] 패턴으로 깔끔히 닫힘.

## 검수 통과 후 TMS 배선
1. 접수완료: submit/route.ts `as_received` → `as_visit_booked` (직접방문 분기) + `change_request_link` 추가.
2. 리마인드 24h/2h: 신규 크론 `api/cron/repair-visit-remind` (D-1·당일 발송, 중복방지 컬럼).
3. 변경/취소: 신규 public API `api/repair/public/resched`·`cancel` (아래 페이지가 호출) → 각각 `as_visit_rescheduled`·`as_visit_cancelled` 발송.
4. make-webhook.ts: `NotifyTemplate` 5종 추가, 라우팅·변수 매핑.
