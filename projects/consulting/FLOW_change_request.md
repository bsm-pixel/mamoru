# 고객 일정 변경/취소 요청 — 전체 흐름도

> 최종 수정: 2026-02-21

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

## 흐름 B-1: 매장방문 — 셀프서비스 (자동취소 + 재예약)

```
1. 고객: "일정확인/변경" 버튼 클릭
      │
2. page_change_request.html 로드 (?uid=UID)
      │  └─ GAS API: getReservationInfo(uid) → type = "매장 방문"
      │
3. 예약 정보 카드 표시 + [일정 변경] [예약 취소] 버튼
      │
4a. [일정 변경] 선택:
      │  └─ 재예약 안내 페이지로 전환
      │     "기존 예약을 취소하고 다시 예약해주세요"
      │     [기존 예약 취소 후 재예약하기] 버튼
      │        │
      │        ├─ GAS: submitChangeRequest(cancel) → 즉시 CANCELLED
      │        └─ 매장방문 접수 페이지(mamoru.kr/52)로 이동
      │           → 고객이 새로 예약 → 즉시 확정 → 새 confirmed 알림톡
      │
4b. [예약 취소] 선택:
      │  └─ 취소 사유 선택 + 메모(선택)
      │  └─ GAS: submitChangeRequest(cancel) → 즉시 CANCELLED
      │  └─ 페이지: "예약이 취소되었습니다" 완료 화면
      │
5. 관리자: 이메일 알림 수신 (TMS에서도 확인 가능)
```

**알림톡 없음** — 매장방문 변경/취소 시 고객에게 별도 알림톡 발송하지 않음

---

## 흐름 B-2: 출장 — 일정변경 요청 → 관리자 재제안

```
1. 고객: "일정확인/변경" 버튼 클릭
      │
2. page_change_request.html 로드 (?uid=UID)
      │  └─ GAS API: getReservationInfo(uid) → type = "출장 요청"
      │
3. 예약 정보 카드 표시 + [일정 변경] [예약 취소] 버튼
      │
4a. [일정 변경] 선택:
      │  └─ "일정 관련 요청사항" 입력 (textarea)
      │  └─ GAS: submitChangeRequest(change) → CHANGE_REQUESTED
      │        │
      │        ├─ (a) 비고 컬럼 append
      │        ├─ (b) 상태 → CHANGE_REQUESTED
      │        ├─ (c) Supabase 동기화
      │        ├─ (d) 관리자 이메일 알림
      │        └─ (e) change_request_received 알림톡
      │              → "출장 일정 변경 요청이 접수되었습니다"
      │
4b. [예약 취소] 선택:
      │  └─ 취소 사유 선택 + 메모(선택)
      │  └─ GAS: submitChangeRequest(cancel) → 즉시 CANCELLED
      │  └─ 페이지: "예약이 취소되었습니다" 완료 화면
      │
5. 관리자: 요청 확인 후 처리
      ├─ 변경 → adminReschedule → 새 일정 제안 알림톡
      └─ 필요 시 전화/톡으로 직접 조율
```

---

## 흐름 C: 관리자 처리 (출장 일정변경)

```
1. TMS 대시보드
      ├─ 출장요청: "대기" 탭에 CHANGE_REQUESTED 표시
      └─ 매장방문: "변경/취소" 탭에 표시 (취소 건만)

2. 관리자: 요청 확인 후 처리
      ├─ 일정 변경 → adminReschedule → RESCHEDULED 알림톡
      └─ 취소 확인 → 이미 CANCELLED 처리됨 (추가 작업 불필요)
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
| CHANGE_REQUEST_RECEIVED | change_request_received | - | 없음 (출장 일정변경 접수 확인) |

---

## change_request_received 템플릿 (출장 전용)

```
[마모루] #{name}님, 출장 일정 변경 요청이 접수되었습니다.

■ 현재 출장 예약
- 예약일: #{date}
- 예약시간: #{time}
- 출장지: #{address}

■ 요청사항
#{request_detail}

확인 후 새로운 일정을 안내드리겠습니다.
영업시간 외 접수 시 다음 영업일에 안내드립니다.
```

변수: `name`, `date`, `time`, `address`, `request_detail`

---

## 파일 위치

| 파일 | 역할 |
|------|------|
| `projects/consulting/Code.gs` | GAS 백엔드 (API + 헬퍼) |
| `projects/consulting/page_change_request.html` | 고객 대면 셀프서비스 페이지 |
| `TMS/.../store-visit-list.tsx` | TMS 매장방문 탭 (변경/취소 탭) |
| `TMS/.../field-request-list.tsx` | TMS 출장요청 탭 (대기에 change_requested 포함) |
| `TMS/.../format.ts` | 상태 라벨/색상 정의 |
