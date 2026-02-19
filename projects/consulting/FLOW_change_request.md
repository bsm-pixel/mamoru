# 고객 일정 변경/취소 요청 — 전체 흐름도

> 최종 수정: 2026-02-19

---

## 시스템 구성 요소

```
[고객 카카오톡] ← 솔라피 ← Make ← GAS (Code.gs)
                                    ↕
                              Google Sheets
                                    ↕
                               Supabase
                                    ↕
                             TMS (React App)
                                    ↕
                            관리자 대시보드
```

---

## 흐름 A: 확정 → 고객에게 변경 버튼 전달

```
1. 관리자: 예약 확정 (매장방문/출장/일정변경)
      │
2. GAS: 확정 페이로드 생성 (change_request_link 포함)
      │  ├─ CONFIRMED          (매장방문 신규)
      │  ├─ CONFIRMED_BY_TOKEN (고객 토큰 확정)
      │  ├─ RESCHEDULED        (관리자 일정변경)
      │  └─ FIELD_CONFIRMED    (출장 확정)
      │
3. GAS → Make (postMake_): payload 전달
      │
4. Make: 솔라피 모듈로 알림톡 발송
      │  └─ 변수 매핑: #{change_request_link}
      │
5. 솔라피: 확정 알림톡 + "일정확인/변경" WL 버튼
      │  └─ 버튼 URL: https://#{change_request_link}
      │     = https://bsm-pixel.github.io/mamoru/projects/consulting/page_change_request.html?uid=UID
      │
6. 고객: 카톡 대화방에서 버튼 상시 접근 가능
```

**리마인드 알림톡 (24H/2H)**: change_request_link **미포함** — 버튼 없음

---

## 흐름 B: 고객 셀프서비스 요청

```
1. 고객: "일정확인/변경" 버튼 클릭
      │
2. page_change_request.html 로드
      │  └─ URL: ?uid=UID
      │
3. JS → GAS API: getReservationInfo(uid)
      │  └─ CONFIRMED/ASSIGNED 상태만 허용
      │
4. 예약 정보 카드 표시
      │  ├─ 고객명, 상담방식, 예약일, 예약시간
      │  └─ [일정 변경] [예약 취소] 버튼
      │
5a. [일정 변경] 선택:
      │  └─ 희망 일시 입력 + 메모(선택)
      │
5b. [예약 취소] 선택:
      │  └─ 취소 사유 선택 + 메모(선택)
      │
6. JS → GAS API: submitChangeRequest(uid, reqType, reason, memo, hopeDate)
```

---

## 흐름 C: 요청 접수 → 관리자 처리

```
1. GAS: submitChangeRequest_ 실행
      │
      ├─ (a) 비고 컬럼 append
      │       [고객 변경요청 MM-dd HH:mm] 희망: ...
      │       [고객 취소요청 MM-dd HH:mm] 사유: ...
      │
      ├─ (b) 상태 → CHANGE_REQUESTED
      │
      ├─ (c) Supabase 동기화 (pushToSupabase_)
      │
      ├─ (d) 관리자 이메일 알림 (bsm@mamoru.kr)
      │
      └─ (e) 접수 확인 알림톡
             GAS → Make (CHANGE_REQUEST_RECEIVED)
                → 솔라피 (change_request_received 템플릿)
                → 고객 카톡: "요청이 접수되었습니다"

2. TMS 대시보드
      ├─ 매장방문: "변경/취소" 탭에 표시
      └─ 출장요청: "대기" 탭에 change_requested 포함

3. 관리자: 요청 확인 후 처리
      ├─ 변경 → adminReschedule → RESCHEDULED 알림톡
      └─ 취소 → 상태 CANCELLED
```

---

## 페이로드 매핑 요약

| 이벤트 | 템플릿 | change_request_link | 버튼 |
|--------|--------|:-------------------:|:----:|
| CONFIRMED | confirmed | O | "일정확인/변경" WL |
| CONFIRMED_BY_TOKEN | confirmed | O | "일정확인/변경" WL |
| RESCHEDULED | rescheduled | O | "일정확인/변경" WL |
| FIELD_CONFIRMED | field_confirmed | O | "일정확인/변경" WL |
| REMINDER_24H | remind24 | X | 없음 |
| REMINDER_2H | remind2 | X | 없음 |
| CHANGE_REQUEST_RECEIVED | change_request_received | - | 없음 (접수 확인만) |

---

## 파일 위치

| 파일 | 역할 |
|------|------|
| `projects/consulting/Code.gs` | GAS 백엔드 (API + 헬퍼) |
| `projects/consulting/page_change_request.html` | 고객 대면 셀프서비스 페이지 |
| `TMS/.../store-visit-list.tsx` | TMS 매장방문 탭 (변경/취소 탭) |
| `TMS/.../field-request-list.tsx` | TMS 출장요청 탭 (대기에 change_requested 포함) |
| `TMS/.../format.ts` | 상태 라벨/색상 정의 |
