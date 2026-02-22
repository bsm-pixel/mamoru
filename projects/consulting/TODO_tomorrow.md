# 할 일 목록 — 2026-02-23 (일)

> 작성: 2026-02-22
> 목표: 솔라피 검수 제출 + TMS AS 기능 구현 착수

---

## 1. 솔라피 알림톡 템플릿 검수 제출

### 1-1. 상담 알림톡 13종 (버튼·메타데이터 확정 완료)
- [ ] confirmed (매장 확정) — WL 일정확인/변경 + BC 1:1 문의하기
- [ ] cancelled (매장 취소) — BC 1:1 문의하기
- [ ] rescheduled (매장 일정변경) — WL 일정확인/변경 + BC 1:1 문의하기
- [ ] remind24 (매장 24h 리마인드) — 퀵버튼 BC 1:1 문의하기
- [ ] remind2 (매장 2h 리마인드) — 퀵버튼 BC 1:1 문의하기
- [ ] suggest (출장 시간 제안) — WL 일정 선택하기 + BC 1:1 문의하기
- [ ] field_confirmed (출장 확정) — WL 일정확인/변경 + BC 1:1 문의하기
- [ ] field_cancelled (출장 취소) — BC 1:1 문의하기
- [ ] field_rescheduled (출장 일정변경) — WL 일정확인/변경 + BC 1:1 문의하기
- [ ] field_remind_24h (출장 24h 리마인드) — 퀵버튼 BC 1:1 문의하기
- [ ] field_remind_2h (출장 2h 리마인드) — 퀵버튼 BC 1:1 문의하기
- [ ] field_delayed (출장 지연 안내) — BC 1:1 문의하기
- [ ] change_request_received (출장 변경요청 접수) — 버튼 없음

> 참고: `참고이미지_스크린샷/솔라피_알림톡_버튼_메타데이터_정리.txt`

### 1-2. 톡상담 알림톡 2종
- [ ] talk_received (톡상담 접수) — BC 1:1 문의하기
- [ ] talk_ready (톡상담 시작) — BC 톡상담 시작하기

### 1-3. AS 알림톡 1종
- [ ] as_received (AS 접수 완료) — WL AS 안내 확인 + BC 1:1 문의하기

### 1-4. 리뷰 유도 알림톡 (초안 작성 필요)
- [ ] 상담 후기 유도 템플릿 초안 작성
- [ ] AS 후기 유도 템플릿 초안 작성
- [ ] 검수 제출

---

## 2. Make 시나리오 분기 연결

### 검수 승인 후 진행 (승인 대기 중이면 2-1부터)
- [ ] 2-1. 기존 4종 (confirmed/cancelled/rescheduled/remind) 변수 매핑 업데이트
- [ ] 2-2. 신규 9종 분기 추가 (Router에서 template 값 기준)
- [ ] 2-3. 톡상담 2종 분기 추가
- [ ] 2-4. AS 1종 분기 추가

---

## 3. TMS AS 기능 구현 착수

### 3-1. AS 안내 페이지 (GitHub Pages)
- [ ] `projects/as/page_as_guide.html` 생성
- [ ] 섹션: 진행과정 / 포장방법 / 비용안내
- [ ] GitHub Pages push + 접근 확인

### 3-2. TMS AS 모듈 (Phase 7 착수)
- [ ] Supabase AS 테이블 설계 (접수→수거→입고→수리→출고)
- [ ] AS 접수 UI 스캐폴드
- [ ] AS 상태 보드 (칸반 or 탭)

> 상세: `memory/as_system_brief.md` 참조

---

## 4. E2E 테스트 계획 (검수 승인 후)

> 원칙: 과정별 순차 테스트 → 합격 → 다음 단계
> 문제 발생 시: 개선 → 재테스트 → 합격 확인 후 다음 진행

### 라운드 1: 매장방문 플로우
```
접수 → 확정 알림톡 → 일정확인/변경 페이지 → 변경(재예약) → 취소
→ 리마인드 24h → 리마인드 2h
```
- [ ] 1-1. 매장 확정 → confirmed 알림톡 수신 + 버튼 확인
- [ ] 1-2. 버튼 → page_change_request.html 정상 로드
- [ ] 1-3. 일정 변경 → 자동취소 + 재예약 안내
- [ ] 1-4. 취소 → 즉시 CANCELLED + 관리자 이메일
- [ ] 1-5. 리마인드 24h/2h → 퀵버튼 확인

### 라운드 2: 출장요청 플로우
```
접수 → 시간제안 → 고객선택/재요청 → 확정 → 변경요청 → 일정변경
→ 리마인드 24h → 리마인드 2h → 지연안내
```
- [ ] 2-1. 출장 접수 확인
- [ ] 2-2. suggest 알림톡 → 캘린더 페이지 정상 로드
- [ ] 2-3. 고객 선택 → field_confirmed 알림톡
- [ ] 2-4. 재요청 → 재제안 사이클
- [ ] 2-5. 일정변경 요청 → change_request_received 알림톡
- [ ] 2-6. 관리자 일정변경 → field_rescheduled 알림톡
- [ ] 2-7. 출장 취소 → field_cancelled 알림톡
- [ ] 2-8. 리마인드 24h/2h → 퀵버튼 확인
- [ ] 2-9. TMS 지연안내 → field_delayed 알림톡

### 라운드 3: 톡상담 플로우
```
접수 → talk_received → 관리자 "상담 시작" → talk_ready → 상담 진행
```
- [ ] 3-1. 톡상담 접수 → talk_received 자동 발송
- [ ] 3-2. TMS "상담 시작" → talk_ready 발송
- [ ] 3-3. 고객 카카오 채널 진입 확인

### 라운드 4: AS 플로우
```
접수 → as_received → AS 안내 페이지 확인
```
- [ ] 4-1. AS 접수 → as_received 알림톡
- [ ] 4-2. "AS 안내 확인" 버튼 → page_as_guide.html 로드

### 라운드 5: 메타데이터 검증
- [ ] 5-1. BC 버튼 클릭 → 해피톡 상담사 화면에 메타데이터 표시 확인
- [ ] 5-2. 메타데이터 값 정확성 (#{type}_#{template}_#{name} → 한글 치환 확인)

---

## 참고 파일
- 버튼/메타데이터 총정리: `참고이미지_스크린샷/솔라피_알림톡_버튼_메타데이터_정리.txt`
- 전체 흐름도: `projects/consulting/FLOW_change_request.md`
- GAS 코드: `projects/consulting/Code.gs`
- TMS 로드맵: `projects/Total_Management_System/docs/TMS_ROADMAP.md`
- AS 브리프: `memory/as_system_brief.md`
