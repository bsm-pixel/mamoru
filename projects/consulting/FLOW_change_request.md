# 고객 일정 변경/취소 요청 — 전체 흐름도

> 최종 수정: 2026-02-22

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

**리마인드 알림톡 (24H/2H)**: change_request_link **미포함** — 1:1 문의하기 퀵버튼만

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
| RESCHEDULED (매장) | rescheduled | O | "일정확인/변경" WL |
| RESCHEDULED (출장) | field_rescheduled | O | "일정확인/변경" WL + 1:1 문의 |
| FIELD_CONFIRMED | field_confirmed | O | "일정확인/변경" WL + 1:1 문의 |
| SUGGESTED_TIMES | suggest | X | "일정 선택하기" WL |
| FIELD_CANCELLED | field_cancelled | X | 1:1 문의하기 (상담톡) |
| FIELD_REMIND_24H | field_remind_24h | X | 1:1 문의하기 (상담톡) |
| FIELD_REMIND_2H | field_remind_2h | X | 1:1 문의하기 (상담톡) |
| FIELD_DELAYED | field_delayed | X | 1:1 문의하기 (상담톡) |
| REMINDER_24H | remind24 | X | 1:1 문의하기 (퀵버튼) |
| REMINDER_2H | remind2 | X | 1:1 문의하기 (퀵버튼) |
| CHANGE_REQUEST_RECEIVED | change_request_received | - | 없음 |

---

## 출장 전용 알림톡 템플릿 (톤: 마침표 없음, 이모지+중점 스타일)

### suggest (출장 시간 제안)
```
#{name}님, 안녕하세요
요청하신 출장 상담 일정을 준비했어요

📅 아래 버튼을 눌러 원하시는 일정을 선택해 주세요

· 제안된 일정이 어려우시면 페이지 하단에서 다른 일정을 요청하실 수 있어요

상담톡 운영시간 : 10:00 ~ 21:00
```
변수: `name`, `confirm_link`
버튼: 일정 선택하기 (WL) → #{confirm_link}

### field_rescheduled (출장 일정 변경 — 관리자 직접 확정)
```
#{name}님, 안녕하세요
조율해 주신 내용으로 출장 상담 일정이 변경되었습니다

✅ 변경된 상담 일정
• 날짜: #{visit_date}
• 시간: #{visit_time}
• 방문 주소: #{address}

🔔 꼭 확인해 주세요
• 원활한 상담을 위해 해당 시간을 미리 비워주시길 부탁드려요
• 일정 변경이나 취소가 필요하시면 아래 버튼을 눌러주세요

상담톡 운영시간 : 10:00 ~ 21:00
```
변수: `name`, `visit_date`, `visit_time`, `address`, `change_request_link`
버튼 1: 일정확인/변경 (WL) → #{change_request_link}
버튼 2: 1:1 문의하기 (상담톡연결)
Make 분기: RESCHEDULED 이벤트에서 `type='출장 요청'`이면 이 템플릿 사용

### field_confirmed (출장 확정)
```
#{name}님, 안녕하세요
요청하신 출장 방문 상담 일정이 확정되었습니다

✅ 확정된 상담 일정
• 확정 날짜: #{visit_date}
• 확정 시간: #{visit_time}
• 방문 주소: #{address}

🔔 꼭 확인해 주세요
• 원활한 상담을 위해 해당 시간을 미리 비워주시길 부탁드려요
• 일정 변경이나 취소가 필요하시면 아래 버튼을 눌러주세요

상담톡 운영시간 : 10:00 ~ 21:00
```
변수: `name`, `visit_date`, `visit_time`, `address`, `change_request_link`
버튼 1: 일정확인/변경 (WL) → #{change_request_link}
버튼 2: 1:1 문의하기 (상담톡연결)

### field_cancelled (출장 취소)
```
#{name}님, 안녕하세요
요청하신 출장 방문 상담이 정상적으로 취소되었습니다

✅ 취소된 상담 내역
• 상담 일정: #{visit_date} #{visit_time}
• 방문 주소: #{address}

· 다음 기회에 더 좋은 서비스로 뵙기를 바라겠습니다
· 새로운 상담이 필요하시면 언제든 편하게 다시 신청해 주세요

상담톡 운영시간 : 10:00 ~ 21:00
```
변수: `name`, `visit_date`, `visit_time`, `address`
버튼: 1:1 문의하기 (상담톡연결)

### field_remind_24h (출장 24시간 전 리마인드)
```
#{name}님, 안녕하세요
내일 출장 상담이 예정되어 있어요

✅ 확정된 상담 일정
• 방문 날짜: #{visit_date} (내일)
• 방문 시간: #{visit_time}
• 방문 주소: #{address}

🔔 꼭 확인해 주세요
• 원활한 상담을 위해 해당 시간을 미리 비워주시길 부탁드려요

상담톡 운영시간 : 10:00 ~ 21:00
```
변수: `name`, `visit_date`, `visit_time`, `address`
버튼: 1:1 문의하기 (상담톡연결)

### field_remind_2h (출장 2시간 전 리마인드)
```
#{name}님, 안녕하세요
오늘 출장 상담이 곧 시작됩니다

✅ 오늘 상담 일정
• 방문 시간: #{visit_time}
• 방문 주소: #{address}

🔔 안내 사항
• 원활한 상담을 위해 약속 시간을 비워주시길 부탁드려요
• 교통 상황에 따라 지연 발생 시 알림톡으로 미리 안내드릴게요

상담톡 운영시간 : 10:00 ~ 21:00
```
변수: `name`, `visit_time`, `address`
버튼: 1:1 문의하기 (상담톡연결)

### field_delayed (출장 방문 지연 안내)
```
#{name}님, 안녕하세요
죄송합니다, 현재 교통 상황으로 인해
약 #{delay_min}분 정도 늦을 예정입니다

✅ 오늘 상담 일정
• 원래 시간: #{visit_time}
• 예상 도착: #{visit_time_revised}
• 방문 주소: #{address}

· 불편을 드려 죄송하며 최대한 빠르게 찾아뵙겠습니다

상담톡 운영시간 : 10:00 ~ 21:00
```
변수: `name`, `delay_min`, `visit_time`, `visit_time_revised`, `address`
버튼: 1:1 문의하기 (상담톡연결)
TMS 구현: 상담 카드 → "지연 안내" 버튼 → 5/10/15/20/30분 선택 → 알림톡 자동 발송

### change_request_received (출장 일정변경 요청 접수)
```
#{name}님, 안녕하세요
출장 상담 일정 변경 요청이 접수되었습니다

· 확인 후 새로운 일정을 안내드릴게요
· 영업시간 외 접수 시 다음 영업일에 안내드립니다

상담톡 운영시간 : 10:00 ~ 21:00
```
변수: `name`
버튼: 없음

---

## 톡상담 흐름

```
1. 고객: 진단 테스트 완료 → 톡상담 접수 (진단 결과 첨부됨)
      │
2. GAS: 자동 알림톡 발송 (talk_received)
      │  └─ "접수 완료, 상담 가능 시 알림톡으로 안내드릴게요"
      │
3. 관리자: TMS 톡상담 탭에서 접수 확인 + 진단 결과 검토
      │  └─ 시간 여유 생길 때 "상담 시작" 버튼 클릭
      │
4. GAS → Make: 알림톡 발송 (talk_ready)
      │  └─ "전문가 준비 완료, 버튼 누르고 '상담시작' 입력해 주세요"
      │
5. 고객: 카카오톡 채널 진입 → "상담시작" 입력
      │
6. 해피톡: 상담 시작 (관리자에게 진단 결과 참조 가능)
```

**상태 흐름**: REQUESTED → READY → (상담 진행) → COMPLETED

## 톡상담 알림톡 템플릿

### talk_received (톡상담 접수 완료)
```
#{name}님, 안녕하세요
미용가위 톡상담 요청이 접수되었습니다

✅ 접수 유형
• 미용가위 톡상담

· 진단 결과를 바탕으로 전문가가 맞춤 상담을 준비해요
· 상담 가능 시 알림톡으로 안내드릴게요

상담톡 운영시간 : 10:00 ~ 21:00
```
변수: `name`
버튼: 1:1 문의하기 (상담톡연결)

### talk_ready (톡상담 시작 — 관리자 트리거)
```
#{name}님, 안녕하세요
톡상담 준비가 완료되었습니다

✅ 진단 결과 확인 완료
• 담당 전문가가 배정되었어요

💬 상담 시작 방법
• 아래 버튼을 누르신 후 "상담시작"이라고 입력해 주세요
• 입력하시면 바로 전문가와 연결됩니다

상담톡 운영시간 : 10:00 ~ 21:00
```
변수: `name`
버튼: 톡상담 시작하기 (상담톡연결)
TMS 구현: 톡상담 탭 → "상담 시작" 버튼 → talk_ready 알림톡 자동 발송 → 상태: REQUESTED → READY

---

## 페이로드 매핑 요약 (톡상담)

| 이벤트 | 템플릿 | 트리거 | 버튼 |
|--------|--------|--------|:----:|
| TALK_RECEIVED | talk_received | 접수 시 자동 | 1:1 문의하기 (상담톡) |
| TALK_READY | talk_ready | 관리자 수동 | 톡상담 시작하기 (상담톡) |

---

## 파일 위치

| 파일 | 역할 |
|------|------|
| `projects/consulting/Code.gs` | GAS 백엔드 (API + 헬퍼) |
| `projects/consulting/page_change_request.html` | 고객 대면 셀프서비스 페이지 |
| `TMS/.../store-visit-list.tsx` | TMS 매장방문 탭 (변경/취소 탭) |
| `TMS/.../field-request-list.tsx` | TMS 출장요청 탭 (대기에 change_requested 포함) |
| `TMS/.../format.ts` | 상태 라벨/색상 정의 |

---

## 최적화 로드맵: Make 제거 → 솔라피 직접 호출

> 상태: 검수 승인 후 전환 예정

### 현재 구조 (Make 경유)
```
GAS/TMS → postMake_(event, payload)
             → Make 웹훅 수신
             → Router: template 값으로 13개 분기
             → 각 분기마다 솔라피 API 모듈 → 알림톡 발송
```

### 전환 후 구조 (솔라피 직접 호출)
```
GAS/TMS → postSolapi_(template, phone, variables)
             → TEMPLATE_MAP에서 솔라피 템플릿ID 매핑
             → 솔라피 REST API 직접 호출 → 알림톡 발송
```

### 전환 시 변경 범위

| 대상 | 변경 내용 |
|------|-----------|
| GAS `postMake_()` | `postSolapi_()` 로 교체 — 솔라피 API Key + 템플릿ID 매핑 |
| TMS `make-webhook.ts` | `sendNotification()` 내부를 솔라피 REST API로 교체 |
| Make 시나리오 | 비활성화 후 삭제 |
| 솔라피 API Key | GAS Script Properties + TMS .env에 추가 |

### 전환 절차
```
1. 솔라피 검수 승인 완료
2. Make 경유로 전체 13종 알림톡 정상 발송 확인
3. GAS에 postSolapi_() 함수 작성 (TEMPLATE_MAP + 솔라피 API 호출)
4. TMS sendNotification() 내부를 솔라피 직접 호출로 교체
5. 테스트 환경에서 13종 전체 발송 테스트
6. postMake_() → postSolapi_() 전환
7. Make 시나리오 비활성화 → 1주일 모니터링 → 삭제
```

### 기대 효과
- Make 월 구독 비용 절감
- 중간 단계 제거로 알림톡 발송 속도 향상
- Make Router 분기 관리 불필요 — 코드 내 TEMPLATE_MAP으로 일원화
