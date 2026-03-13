# MAMORU TODO

> 단일 파일 누적 관리 — 완료 시 [x] 체크, 신규 항목 추가
> 최종 수정: 2026-03-13

---

## 🔧 진행중

### Make Router / 알림톡
- [x] Make Router 상담 17종 분기 연결 ✅ 03-12
- [x] Make Router 복원수리 6종 분기 연결 ✅ 03-12
- [x] Make Router 계약서 1종 분기 연결 ✅ 03-12
- [ ] Make Router **리뷰 작성 요청 템플릿** 연결 (후기 템플릿)
- [ ] 솔라피 BC 메타데이터 제거 후 재검수 대기 중
- [ ] 솔라피 confirmed/rescheduled/field_confirmed 템플릿에 WL 버튼(일정확인/변경) 추가 → 재검수
- [ ] Make: 확정 3개 시나리오에 change_request_link 변수 매핑

### 고객 대면 페이지 (page_guide.html)
- [x] Before & After 캐러셀 구조 구현 ✅ 03-12
- [x] 슬라이드 1~3 이미지 삽입 ✅ 03-12
- [x] 전후영상 2, 3 삽입 ✅ 03-12
- [ ] Before & After 슬라이드 4, 5 이미지 삽입
- [ ] 전후영상 4, 5 촬영/삽입
- [ ] 작업과정 타임라인 이미지 9장 (사전검수~스토퍼장착)

---

## 📋 대기

### UI 동작 검증 (수동 — Phase ERP 전 모듈)
- [ ] /sales/new 판매입력 (딜러 고객 시 도매가 자동 적용)
- [ ] /contracts/new 전자문서 (제품 모달 + 서명 2개 + 수령방법 + 선납/잔금)
- [ ] /products/[id]/serials 시리얼 등록
- [ ] /customers 고객 목록/상세
- [ ] /products/new 제품 등록 (3단 가격)
- [ ] /purchasing/new 발주 작성 → 입고 → 재고 증가
- [ ] /inventory 재고 현황 + 저재고 필터
- [ ] /reports 회계 리포트 + 엑셀 다운로드
- [ ] /reports/transaction 거래내역서 인쇄

### 시스템 기능
- [ ] Phase 3: 아임웹 자동 동기화 (Vercel Cron 5~10분)
- [ ] Phase 7: 주소 수정 시 다음 주소검색 API 연동
- [ ] Phase 7: 사진 업로드 Supabase Storage 연동
- [ ] Phase 7: 수리내역서 자동 생성 (Before/After 타임라인)
- [ ] Phase 6: 계약서 PDF/이미지 생성
- [ ] 재검수 승인 후 E2E 테스트 (상담+복원수리 전체 플로우)

### 제품구매 리뷰 시스템 (03-13 구현 완료, 후속 작업)
- [x] DB 마이그레이션 011 — orders.review_requested_at ✅ 03-13
- [x] 리뷰 API (info/submit) purchase 타입 추가 ✅ 03-13
- [x] 리뷰 폼 purchase 타입 + 제품 선택 UI ✅ 03-13
- [x] 리뷰 보드 제품 썸네일/제품명 표시 ✅ 03-13
- [x] TMS sync 배송완료 자동 감지 + 알림톡 트리거 ✅ 03-13
- [x] 수동 리뷰 요청 API (/api/imweb/orders/[id]/review-request) ✅ 03-13
- [ ] 솔라피 `purchase_review_request` 템플릿 등록 → 카카오 검수 대기
- [ ] Make 시나리오 `PURCHASE_REVIEW_REQUEST` 분기 추가
- [ ] 1:1 문의 버튼 URL 확정 (카카오 채널 채팅 링크)
- [ ] Vercel 배포 후 E2E 테스트 (수동 버튼 → 알림톡 → 리뷰 폼 → 제출)

### 아임웹 알림 설정
- [ ] 이메일 발송 설정 (결제완료/발송안내/취소완료/환불완료 등 ON)
- [ ] 알림톡 발송 설정 확인 (결제완료/발송안내/취소완료/환불완료 ON, 배송준비/배송완료 OFF)

### 미착수
- [ ] 통합 리뷰 시스템 (memory/REVIEW_SYSTEM_BRIEF.md 참조)
