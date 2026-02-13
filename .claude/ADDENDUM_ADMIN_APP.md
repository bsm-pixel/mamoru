# ADDENDUM_ADMIN_APP.md - Admin / Backoffice App Rules

## 0. 목표
- 운영 실수 방지(가드레일) + 권한 통제 + 변경 이력(감사 로그) 확보

---

## 1. UX 원칙(관리자용)
- “빠름”보다 “실수 방지” 우선:
  - 파괴적 작업(삭제/완료/취소/환불/상태 확정)은 확인 단계 + 되돌리기(undo) 경로 제공
- 리스트/상세 화면은 상태값(ENUM)과 다음 액션이 즉시 보이게 구성
- 에러 메시지는 해결 방법을 포함

---

## 2. 권한/역할
- 역할(Role) 기준으로 화면/액션/데이터 접근을 제한
- 민감정보는 최소 노출 + 마스킹 기본값

---

## 3. 데이터 무결성 & 상태 모델
- 상태값(ENUM)은 문서로 고정하고, 전이(transition) 규칙을 명확히 한다.
- 변경 이력:
  - 누가/언제/무엇을/왜 변경했는지 로그를 남긴다.

---

## 4. AppSheet 네이티브 우선(강화)
- 관리 자동화(상태 변경/알림/문서/차단)는 AppSheet Automation/Action/Expression 우선
- 구현 시 AppSheet 에디터 경로를 명시:
  - 예) Automation > New Bot > Event ... / Action > New action ...