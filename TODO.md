# MAMORU 시스템 구축 — TODO

> 하위 프로젝트 5개로 분류 — 완료 시 [x] 체크, 신규 항목 추가
> 최종 수정: 2026-03-18

---

## 1. TMS 앱 개발
> TMS 자체 UI / API / DB 기능

### UI 동작 검증 (수동)
- [ ] /sales/new 판매입력 (딜러 고객 시 도매가 자동 적용)
- [ ] /contracts/new 전자문서 (제품 모달 + 서명 2개 + 수령방법 + 선납/잔금)
- [ ] /products/[id]/serials 시리얼 등록
- [ ] /customers 고객 목록/상세
- [ ] /products/new 제품 등록 (3단 가격)
- [ ] /purchasing/new 발주 작성 → 입고 → 재고 증가
- [ ] /inventory 재고 현황 + 저재고 필터
- [ ] /reports 회계 리포트 + 엑셀 다운로드
- [ ] /reports/transaction 거래내역서 인쇄

### 완료 (03-16)
- [x] staleTime 30s 명시 (7개 목록 훅) ✅
- [x] 캐시 무효화 스코프 축소 (hub-stats 자연 갱신 위임) ✅
- [x] 페이지네이션 전체 건수 표시 ✅
- [x] EmptyState 공통 컴포넌트 (4+2개 페이지) ✅
- [x] SearchInput / Pagination 공통 컴포넌트 (6개 페이지) ✅
- [x] Optimistic Update (repair/consultation 상태변경) ✅
- [x] syncOrders upsert 트랜잭션 보호 ✅
- [x] 상담→계약서→판매 CTA 연결 ✅
- [x] useHubStats RPC 최적화 (14쿼리→1 RPC) ✅
- [x] 액션칩 모바일 줄바꿈 정리 ✅
- [x] terracotta → 모노크롬 팔레트 전환 (globals.css) ✅
- [x] repair-detail-panel 높이 flex-1 전환 ✅

### 기능 개발
- [ ] 계약서 PDF/이미지 생성 (html2canvas or 서버사이드)
- [ ] 복원수리: 주소 수정 시 다음 주소검색 API 연동
- [ ] 복원수리: 사진 업로드 Supabase Storage 연동
- [ ] 복원수리: 수리내역서 자동 생성 (Before/After 타임라인)
- [ ] 복원수리: 사진 마킹 (photo-marker.tsx — html2canvas)

---

## 2. 알림톡 · Make 자동화
> 솔라피 템플릿 + Make 시나리오 + 카카오 검수

### 솔라피 재검수
- [ ] BC 메타데이터 제거 후 재검수 대기 중
- [ ] confirmed 템플릿에 WL 버튼(일정확인/변경) 추가 → 재검수
- [ ] rescheduled 템플릿에 WL 버튼 추가 → 재검수
- [ ] field_confirmed 템플릿에 WL 버튼 추가 → 재검수
- [x] `purchase_review_request` 템플릿 등록 → 카카오 검수 승인 완료 ✅ 03-18

### Make Router
- [x] 상담 17종 분기 연결 ✅ 03-12
- [x] 복원수리 6종 분기 연결 ✅ 03-12
- [x] 계약서 1종 분기 연결 ✅ 03-12
- [x] **리뷰 작성 요청 템플릿** 연결 (상담/복원수리 후기) ✅
- [ ] 확정 3개 시나리오에 change_request_link 변수 매핑
- [ ] CHANGE_REQUEST_RECEIVED 이벤트 분기 + 솔라피 모듈 연결
- [x] `PURCHASE_REVIEW_REQUEST` 분기 추가 + 솔라피 템플릿 연결 ✅ 03-18

### 검증
- [ ] 재검수 승인 후 E2E 테스트 (상담 17종 + 복원수리 6종 + 계약서 1종)
- [ ] 검수 안정화 후 Make → 솔라피 직접 호출 전환 검토 (비용 절감)

---

## 3. 고객 페이지 (아임웹 + GitHub Pages)
> 고객이 직접 보는 모든 페이지

### page_guide.html (복원수리 안내)
- [x] Before & After 캐러셀 구조 구현 ✅ 03-12
- [x] 슬라이드 1~3 이미지 삽입 ✅ 03-12
- [x] 전후영상 2, 3 삽입 ✅ 03-12
- [ ] Before & After 슬라이드 4, 5 이미지 삽입
- [ ] 전후영상 4, 5 촬영/삽입
- [ ] 작업과정 타임라인 이미지 9장 (사전검수~스토퍼장착)

### 리뷰 페이지 (03-18 완료)
- [x] page_review.html: 브랜드 가이드 모노크롬 전면 수정 ✅ 03-18
- [x] page_review.html: 카테고리별 속성 태그 선택 UI (컨설팅7/복원수리7/제품구매6 범용) ✅ 03-18
- [x] page_review.html: 제품구매 연속 작성 ("다른 제품 후기도 작성하기" 버튼) ✅ 03-18
- [x] page_reviews.html + ImwebWidgetCode: 태그 칩 표시 (카드+모달) ✅ 03-18
- [x] submit API: tags 배열 수신 → meta.tags 저장 ✅ 03-18

### 메인 페이지 (03-18 진행)
- [x] 섹션3: 복원수리/컨설팅 중복 카드 제거 ✅ 03-18
- [x] 섹션3: 이벤트→브랜드소개 카드 교체 + 클리어런스 배너 추가 ✅ 03-18
- [x] 후기 섹션: 하드코딩 더미→API 실데이터 연동 ✅ 03-18
- [x] 듀얼CTA: picture 태그 PC/모바일 이미지 분리 (이미지 삽입 대기) ✅ 03-18
- [x] 폰트: Pretendard Variable 모바일 본문 최우선 ✅ 03-18
- [ ] 듀얼CTA: 복원수리/컨설팅 PC+모바일 이미지 삽입 (수동)
- [ ] 아임웹 상품진열 공통 CSS 브랜드 가이드 톤 수정

### 기타 고객 페이지
- (신규 작업 발생 시 여기에 추가)

---

## 4. 주문 · 배송 연동
> 아임웹 API + 롯데택배 ALPS + Vercel Cron

### 🚨 롯데택배 SMS 오발송 수정 (03-17 ~ 03-19)
> 롯데택배 SMS에 테스트 데이터(하우하홍, AS 출고, 안성시 주소)가 표시되는 문제
- [ ] GAS Script Properties `LOTTE_CURRENT11` 현재값 확인
- [ ] 테스트 기간에 사용한 송장번호 범위 파악 (ALPS 이력 조회 or 로그 확인)
- [ ] `LOTTE_CURRENT11`을 테스트 미사용 구간으로 건너뛰기 (또는 롯데택배에 테스트 대역 초기화 요청)
- [ ] 롯데택배 API 담당자에 문의: 동일 송장번호 재접수 시 SMS 데이터 갱신 여부 확인
- [ ] 수정 후 실제 주문 건으로 SMS 정상 발송 검증

### 자동 동기화
- [x] Vercel Cron 설정 (오전9시/오후9시 KST) ✅ 03-18
- [x] sync-orders: 아임웹 배송완료 감지 → TMS delivered 전환 ✅ 03-18
- [x] shipping 보호 상태에서 delivered 전환 차단 버그 수정 ✅ 03-18
- [x] 배송완료 시 리뷰 알림톡 자동 발송 트리거 ✅ 03-18
- [ ] 증분 동기화 최적화 (마지막 sync 이후 변경분만)
- [ ] 동기화 상태 대시보드 표시 (마지막 시간, 에러 로그)

### 아임웹 알림 설정
- [ ] 이메일 발송 설정 (결제완료/발송안내/취소완료/환불완료 등 ON)
- [ ] 알림톡 발송 설정 확인 (결제완료/발송안내/취소완료/환불완료 ON, 배송준비/배송완료 OFF)

---

## 5. 리뷰 시스템
> 제품구매 리뷰 + 통합 리뷰 파이프라인

### 제품구매 리뷰 (03-13 코드 완료, 후속 작업)
- [x] DB 마이그레이션 011 — orders.review_requested_at ✅ 03-13
- [x] 리뷰 API (info/submit) purchase 타입 추가 ✅ 03-13
- [x] 리뷰 폼 purchase 타입 + 제품 선택 UI ✅ 03-13
- [x] 리뷰 보드 제품 썸네일/제품명 표시 ✅ 03-13
- [x] TMS sync 배송완료 자동 감지 + 알림톡 트리거 ✅ 03-13
- [x] 수동 리뷰 요청 API ✅ 03-13
- [ ] 1:1 문의 버튼 URL 확정 (카카오 채널 채팅 링크)
- [ ] Vercel 배포 후 E2E 테스트 (수동 버튼 → 알림톡 → 리뷰 폼 → 제출)

### 통합 리뷰 시스템 (미착수)
- [ ] 설계 및 구현 (memory/REVIEW_SYSTEM_BRIEF.md 참조)

---

## 범례
- [x] 완료
- [ ] 미완료
- 솔라피/Make 관련 항목은 **2. 알림톡** 에 통합 (리뷰 템플릿 포함)
