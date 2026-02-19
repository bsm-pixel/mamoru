# 내일 할 일 — 고객 변경/취소 요청 시스템 마무리

> 작성: 2026-02-19
> 상태: GAS 코드 완료, 솔라피/Make 연동 미완료

---

## 1. 솔라피 템플릿 작업

### 1-1. confirmed 템플릿 수정
- [ ] 기존 confirmed 템플릿 열기
- [ ] "일정확인/변경" **WL(웹링크)** 버튼 추가
  - 버튼명: `일정확인/변경`
  - 버튼타입: WL (웹링크)
  - URL: `https://#{change_request_link}`
- [ ] 검수 요청 제출
- [ ] 검수 승인 확인

### 1-2. rescheduled 템플릿 수정
- [ ] 기존 rescheduled 템플릿 열기
- [ ] "일정확인/변경" WL 버튼 추가 (위와 동일)
- [ ] 검수 요청 → 승인 확인

### 1-3. field_confirmed 템플릿 수정
- [ ] 기존 field_confirmed 템플릿 열기
- [ ] "일정확인/변경" WL 버튼 추가 (위와 동일)
- [ ] 검수 요청 → 승인 확인

### 1-4. change_request_received 신규 템플릿 등록
- [ ] 템플릿명: `change_request_received`
- [ ] 내용 초안:
  ```
  [마모루] #{name}님, #{request_type_label} 요청이 접수되었습니다.

  ■ 현재 예약 정보
  - 상담방식: #{type}
  - 예약일: #{date}
  - 예약시간: #{time}

  ■ 요청 내용
  #{request_summary}

  담당자 확인 후 연락드리겠습니다.
  영업시간 외 접수 시 다음 영업일에 처리됩니다.
  ```
- [ ] 변수: name, type, date, time, request_type_label, request_summary
- [ ] 버튼: 없음 (접수 확인 전용)
- [ ] 검수 요청 → 승인 확인

---

## 2. Make 시나리오 작업

### 2-1. 확정 시나리오 3개 — change_request_link 변수 매핑
- [ ] CONFIRMED 시나리오: 솔라피 모듈에 `change_request_link` 변수 추가
- [ ] RESCHEDULED 시나리오: 솔라피 모듈에 `change_request_link` 변수 추가
- [ ] FIELD_CONFIRMED 시나리오: 솔라피 모듈에 `change_request_link` 변수 추가

### 2-2. CHANGE_REQUEST_RECEIVED 신규 분기
- [ ] Webhook 라우터에 `CHANGE_REQUEST_RECEIVED` 이벤트 분기 추가
- [ ] 솔라피 알림톡 모듈 연결 (change_request_received 템플릿)
- [ ] 변수 매핑: name, phone, type, date, time, request_type_label, request_summary

---

## 3. GAS 배포

- [ ] Code.gs clasp push (change_request_link 변경사항 반영)
- [ ] 새 배포 버전 생성

---

## 4. 통합 테스트

### 4-1. 확정 알림톡 테스트
- [ ] 매장방문 확정 → 알림톡에 "일정확인/변경" 버튼 표시 확인
- [ ] 토큰 확정 → 버튼 표시 확인
- [ ] 일정변경 → 버튼 표시 확인
- [ ] 출장 확정 → 버튼 표시 확인
- [ ] 버튼 클릭 → page_change_request.html 정상 로드

### 4-2. 리마인드 알림톡 테스트
- [ ] 24H 리마인드 → 버튼 **없음** 확인
- [ ] 2H 리마인드 → 버튼 **없음** 확인

### 4-3. 셀프서비스 E2E
- [ ] 변경 요청 → 접수 알림톡 수신 확인
- [ ] 취소 요청 → 접수 알림톡 수신 확인
- [ ] TMS 대시보드에 "변경/취소" 탭 표시 확인
- [ ] 관리자 이메일 수신 확인

---

## 참고
- 전체 흐름도: `projects/consulting/FLOW_change_request.md`
- GAS 코드: `projects/consulting/Code.gs`
- 고객 페이지: `projects/consulting/page_change_request.html`
