# 내일 할 일 — 고객 변경/취소 요청 시스템 마무리

> 작성: 2026-02-19
> 수정: 2026-02-21 (매장방문/출장 분기 반영)
> 상태: GAS 코드 + 고객 페이지 완료, 솔라피/Make 연동 미완료

---

## 변경 요약 (2026-02-21)

- **매장방문 일정변경** → 기존 예약 자동취소 + 접수 페이지 재예약 유도 (알림톡 없음)
- **출장 일정변경** → 요청사항 입력 → `change_request_received` 알림톡 발송
- **취소 (공통)** → 즉시 `CANCELLED` 처리, 페이지에서 완료 안내 (알림톡 없음)
- `change_request_received` 템플릿: **출장 일정변경 전용**으로 축소

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

### 1-4. change_request_received 신규 템플릿 등록 (출장 일정변경 전용)
- [ ] 템플릿명: `change_request_received`
- [ ] 내용 초안:
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
- [ ] 변수: name, date, time, address, request_detail
- [ ] 버튼: 없음 (접수 확인 전용)
- [ ] 검수 요청 → 승인 확인

---

## 2. Make 시나리오 작업

### 2-1. 확정 시나리오 3개 — change_request_link 변수 매핑
- [ ] CONFIRMED 시나리오: 솔라피 모듈에 `change_request_link` 변수 추가
- [ ] RESCHEDULED 시나리오: 솔라피 모듈에 `change_request_link` 변수 추가
- [ ] FIELD_CONFIRMED 시나리오: 솔라피 모듈에 `change_request_link` 변수 추가

### 2-2. CHANGE_REQUEST_RECEIVED 신규 분기 (출장 일정변경 전용)
- [ ] Webhook 라우터에 `CHANGE_REQUEST_RECEIVED` 이벤트 분기 추가
- [ ] 솔라피 알림톡 모듈 연결 (change_request_received 템플릿)
- [ ] 변수 매핑: name, phone, date, time, address, request_detail

---

## 3. GAS 배포

- [ ] Code.gs clasp push (매장/출장 분기 + 취소 즉시처리 반영)
- [ ] 배포 업데이트 (새 배포 X)

---

## 4. GitHub Pages 배포

- [ ] page_change_request.html push (매장방문 재예약 + 출장 요청사항 UI)

---

## 5. 통합 테스트

### 5-1. 확정 알림톡 테스트
- [ ] 매장방문 확정 → 알림톡에 "일정확인/변경" 버튼 표시 확인
- [ ] 토큰 확정 → 버튼 표시 확인
- [ ] 일정변경(rescheduled) → 버튼 표시 확인
- [ ] 출장 확정 → 버튼 표시 확인
- [ ] 버튼 클릭 → page_change_request.html 정상 로드

### 5-2. 리마인드 알림톡 테스트
- [ ] 24H 리마인드 → 버튼 **없음** 확인
- [ ] 2H 리마인드 → 버튼 **없음** 확인

### 5-3. 매장방문 셀프서비스 E2E
- [ ] 일정변경 선택 → 재예약 안내 페이지 표시
- [ ] "기존 예약 취소 후 재예약하기" 클릭 → 기존 건 CANCELLED + 접수 페이지 이동
- [ ] 접수 페이지에서 재예약 → 새 confirmed 알림톡 수신
- [ ] 취소 선택 → 즉시 취소 완료 화면 + TMS 상태 CANCELLED 확인
- [ ] 관리자 이메일 수신 확인

### 5-4. 출장 셀프서비스 E2E
- [ ] 일정변경 선택 → 요청사항 입력 폼 표시
- [ ] 제출 → `change_request_received` 알림톡 수신 확인
- [ ] TMS 대시보드에 CHANGE_REQUESTED 상태 표시 확인
- [ ] 취소 선택 → 즉시 취소 완료 화면 + TMS 상태 CANCELLED 확인
- [ ] 관리자 이메일 수신 확인

---

## 6. AS 알림톡 안내 페이지 (GitHub Pages)

> 현재 구글 슬라이드로 만들어둔 AS 안내를 GitHub Pages 단일 페이지로 이전

### 확정 링크 (솔라피 템플릿에 먼저 삽입)
- **AS 안내**: `https://bsm-pixel.github.io/mamoru/projects/as/page_as_guide.html`
- **1:1 문의하기**: 기존 카카오 채널 상담 링크

### 페이지 구성 (page_as_guide.html)
- [ ] 섹션 1: 진행과정 — 접수→수거→입고→수리→출고 단계별 안내
- [ ] 섹션 2: 포장방법 — 이미지+텍스트 가이드
- [ ] 섹션 3: 비용안내 — AS/복원 가격표
- [ ] 파일 생성: `projects/as/page_as_guide.html`
- [ ] GitHub Pages push 후 접근 확인

### 버튼 구성 (AS 관련 모든 알림톡 공통)
- 버튼 1: `AS 안내` (WL) → `https://bsm-pixel.github.io/mamoru/projects/as/page_as_guide.html`
- 버튼 2: `1:1 문의하기` (WL) → 카카오 채널 상담 링크

---

## 참고
- 전체 흐름도: `projects/consulting/FLOW_change_request.md`
- GAS 코드: `projects/consulting/Code.gs`
- 고객 페이지: `projects/consulting/page_change_request.html`
- AS 안내 원본: https://docs.google.com/presentation/d/1gt4tW0pa1apbTc8g8CPFUQ4ZH4xkZ9zTy5Z9-Qkklcg/edit
