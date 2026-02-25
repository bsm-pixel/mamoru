# TMS (Total Management System) 전체 작업 로드맵

> 최종 목적: 마모루 운영의 주문·배송·수리·재고·알림을 하나의 시스템에서 관리
> 최종 수정: 2026-02-25 (Phase 2 대시보드 허브+카테고리 리팩토링)

---

## Phase 1: 주문·배송 코어 ✅ 완료

**목적:** 아임웹 주문을 TMS로 가져오고, 롯데택배(ALPS) 송장을 생성/취소/추적하는 기본 파이프라인 구축

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 1-1 | 아임웹 → TMS 주문 동기화 | ✅ | 아임웹 v2 API 폴링 → Supabase upsert |
| 1-2 | 주문 목록/상세 UI | ✅ | 상태 탭 필터, 검색, 페이지네이션 |
| 1-3 | 송장 생성 (ALPS 접수) | ✅ | 롯데택배 apiSndOut → 12자리 운송장 생성 |
| 1-4 | 송장 취소 (소프트 취소) | ✅ | TMS cancel_pending → ALPS 수동 집하취소 → 자동 감지 |
| 1-5 | 아임웹 송장 연동 | ✅ | 배송대기(STANDBY) 상태에서 PATCH invoice 입력 |
| 1-6 | 동기화 보호 | ✅ | cancel_pending/shipping 상태 덮어쓰기 방지 |

### 운영 플로우 (확정)
```
주문 접수 → 아임웹 "배송대기 처리" (수동 1회)
         → TMS 송장 생성 (ALPS 접수 + 아임웹 자동 연동)
         → 배송 추적

취소 시 → TMS 송장취소 (cancel_pending)
       → ALPS 집하취소 (수동)
       → TMS 자동 감지 → cancelled
       → 아임웹 고객취소 or 관리자 취소 (수동)
```

### 제약사항
- 아임웹 v2 API: 읽기 + 송장입력만 가능 (상태변경 code -99)
- ALPS: 취소 API 없음 → 소프트 취소 패턴
- 고객 취소: 상품준비중까지만 가능 (배송대기 이후 관리자 직접)

---

## Phase 1.5: 상담 취소 연동 ✅ 완료

**목적:** TMS 상담 취소 시 구글 캘린더 삭제 + 아임웹 슬롯 해제 + 알림톡 자동 발송

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 1.5-1 | TMS → GAS 취소 연동 | ✅ | cancelViaGAS GET 호출 → 캘린더 삭제 + 슬롯 해제 |
| 1.5-2 | GAS doGet cancelConsultation 핸들러 | ✅ | POST body 제한 → GET 쿼리 파라미터 방식 전환 |
| 1.5-3 | 취소 알림톡 자동 발송 | ✅ | Make webhook → Solapi 취소 알림톡 |
| 1.5-4 | 취소 확인 모달 (안전장치) | ✅ | 3개 상담유형 모두 확인 모달 거쳐 취소 |
| 1.5-5 | 백그라운드 처리 (빠른 응답) | ✅ | after() API로 응답 즉시 반환, 후속작업 비동기 |

### 운영 플로우 (확정)
```
취소 시 → TMS 취소 버튼 → 확인 모달 ("정말 취소하시겠습니까?")
       → 취소 확정 → Supabase 상태 변경 → UI 즉시 반영
       → (백그라운드) GAS 캘린더 삭제 + 아임웹 슬롯 해제
       → (백그라운드) 알림톡 자동 발송
```

### 해결한 이슈
- GAS 웹앱 외부 POST body 수신 불가 → GET 쿼리 파라미터 전환
- Vercel 환경변수 `\n` 문제 → 재설정
- fire-and-forget → after() 백그라운드 처리 (Vercel 실행 컨텍스트 보장)

---

## Phase 1.6: 고객 일정 변경/취소 요청 ✅ 완료

**목적:** 확정된 예약(매장방문/출장)에 대해 고객이 직접 일정 변경 또는 취소를 요청 — 알림톡 → 셀프서비스 폼 → TMS 연동

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 1.6-1 | page_change_request.html | ✅ | GitHub Pages 고객 대면 페이지 (예약조회+변경/취소 폼) |
| 1.6-2 | GAS API: getReservationInfo | ✅ | uid로 예약 정보 조회 (CONFIRMED/ASSIGNED만) |
| 1.6-3 | GAS API: submitChangeRequest | ✅ | 매장/출장 분기 처리 (아래 참조) |
| 1.6-4 | 확정 알림톡 change_request_link 배치 | ✅ | CONFIRMED/CONFIRMED_BY_TOKEN/RESCHEDULED/FIELD_CONFIRMED에 추가, REMINDER에서 제거 |
| 1.6-5 | TMS 상태/UI 연동 | ✅ | change_requested 타입·라벨·색상·전이·탭·상세카드 |

### 매장방문 / 출장 분기 (2026-02-21 확정)

**매장방문 일정변경:**
```
확정 알림톡 → "일정확인/변경" 버튼 → 셀프서비스 페이지
  → [일정 변경] 선택 → "기존 예약 취소 후 재예약" 안내
  → 버튼 클릭 → adminCancel(skipNotify) 자동취소 → 접수페이지(mamoru.kr/52) 이동
  → 고객이 새로 예약 → 즉시 확정 → 새 confirmed 알림톡
```
- 관리자 개입 없음, 알림톡 없음 (페이지가 피드백)

**출장 일정변경:**
```
확정 알림톡 → "일정확인/변경" 버튼 → 셀프서비스 페이지
  → [일정 변경] 선택 → 요청사항 입력 → 제출
  → GAS: CHANGE_REQUESTED + Supabase + 이메일 + change_request_received 알림톡
  → TMS 출장 "처리대기" 탭 → 관리자 "새 시간 제안"
```

**취소 (공통):**
```
셀프서비스 페이지 → [예약 취소] → 사유 선택 → 제출
  → adminCancel(skipNotify) 즉시 CANCELLED (캘린더 삭제 + 슬롯 해제)
  → 페이지: "예약이 취소되었습니다" + 관리자 이메일
  → 알림톡 미발송 (관리자 수동 취소 시에만 알림톡 발송)
```

### TMS UI 변경 (2026-02-21)
- 매장방문: `변경/취소` 탭 제거 (자동취소+재예약이므로 불필요)
- 출장: `처리대기` 탭에 change_requested 포함, 버튼 "새 시간 제안"
- 상세 페이지: 고객 요청 카드(주황색) — 비고에서 `[고객 변경요청]` 파싱 표시
- 상태 전이: change_requested → suggested/confirmed/on_hold/cancelled

### 수동 작업 필요 (솔라피/Make) — Phase 1.6-6
- [ ] 솔라피: confirmed 템플릿에 "일정확인/변경" WL 버튼 추가 → 재검수
- [ ] 솔라피: rescheduled 템플릿에 "일정확인/변경" WL 버튼 추가 → 재검수
- [ ] 솔라피: field_confirmed 템플릿에 "일정확인/변경" WL 버튼 추가 → 재검수
- [ ] 솔라피: change_request_received 신규 템플릿 등록 (출장 전용, 변수: name/date/time/address/request_detail) → 검수
- [ ] Make: 확정 3개 시나리오에 change_request_link 변수 매핑
- [ ] Make: CHANGE_REQUEST_RECEIVED 이벤트 분기 + 솔라피 모듈 연결 (변수: name/phone/date/time/address/request_detail)

---

## Phase 1.7: 출장 일정 제안 캘린더 UI ✅ 완료

**목적:** 출장 상담 일정 제안 페이지를 텍스트 버튼 나열 → 캘린더 + 라디오 카드 + 재요청 폼 통합 UI로 업그레이드

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 1.7-1 | page_suggest.html 캘린더 UI | ✅ | 달력(제안 날짜 강조) + 라디오 카드 + 확정 모달 |
| 1.7-2 | page_suggest.html 재요청 통합 | ✅ | 하단 토글 → textarea(reason) + 재요청 모달 |
| 1.7-3 | page_reschedule.html 리다이렉트 | ✅ | 기존 알림톡 링크 호환, page_suggest.html로 자동 이동 |
| 1.7-4 | Code.gs reason 파라미터 | ✅ | markResched에 reason 읽기 → 이메일+비고 컬럼 반영 |
| 1.7-5 | GAS 배포 + GitHub Pages 배포 | ✅ | clasp push @285 + Pages 서빙 확인 |
| 1.7-6 | 실 환경 테스트 | 📋 | 실 토큰 테스트, 카카오 인앱 확인 |

---

## Phase 1.8: 알림톡 13종 출장/매장/톡상담 분기 구현 ✅ 완료

**목적:** 기존 매장방문 전용이던 알림톡을 출장/매장/톡상담 유형별로 분기 — GAS 6개 패치 + TMS 7개 파일 변경

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 1.8-1 | GAS adminFieldDelay 신규 함수 | ✅ | 출장 지연 안내: visit_time + delayMin → visit_time_revised 계산 |
| 1.8-2 | GAS doGet fieldDelay/talkReady 액션 | ✅ | TMS_SYNC_KEY 인증, TMS→GAS 호출 엔드포인트 |
| 1.8-3 | GAS submitConsultation 톡상담 알림 | ✅ | type='톡상담' 시 TALK_RECEIVED 자동 발송 |
| 1.8-4 | GAS adminCancel 출장/매장 분기 | ✅ | 출장→field_cancelled, 매장→cancelled |
| 1.8-5 | GAS adminReschedule 출장/매장 분기 | ✅ | 출장→field_rescheduled, 매장→rescheduled |
| 1.8-6 | GAS sendReminders_ 출장/매장 분기 | ✅ | 출장→field_remind_24h/2h, 매장→remind24/2 |
| 1.8-7 | TMS make-webhook.ts 타입 확장 | ✅ | NotifyTemplate 7종 추가 + TEMPLATE_EVENT_MAP |
| 1.8-8 | TMS delay API route 생성 | ✅ | POST /api/consultation/delay → GAS fieldDelay 호출 |
| 1.8-9 | TMS hooks 추가 | ✅ | useFieldDelay + useStartTalkConsult |
| 1.8-10 | TMS [id]/route.ts 자동 알림 분기 | ✅ | getAutoNotifyTemplate() — 상담유형별 템플릿 선택 |
| 1.8-11 | TMS talk-consult-list.tsx 상담시작 | ✅ | "상담 시작" → useStartTalkConsult → talk_ready 발송 |
| 1.8-12 | TMS field-request-list.tsx 지연안내 | ✅ | "지연 안내" 버튼 + DelaySelectModal (5~30분) |
| 1.8-13 | 고객 페이지 경고 문구 추가 | ✅ | 출장 변경/취소 시 스케줄 조율 안내 (change_request + suggest) |
| 1.8-14 | 솔라피 버튼·메타데이터 총정리 | ✅ | 17개 템플릿 BC/WL/퀵버튼 + 메타데이터 #{type}_#{template}_#{name} |
| 1.8-15 | GAS 배포 @286 + GitHub Pages 배포 | ✅ | clasp push + git push |

### 알림톡 템플릿 17종 체계
```
매장방문 (5종): confirmed, cancelled, rescheduled, remind24, remind2
출장요청 (9종): request, suggest, field_confirmed, field_cancelled,
                field_rescheduled, field_remind_24h, field_remind_2h,
                field_delayed, change_request_received
톡상담   (2종): talk_received, talk_ready
복원수리 (1종): as_received
```

### 버튼·메타데이터 설계
- WL 버튼: 6개 (일정확인/변경, 일정 선택하기, 복원수리 안내 확인)
- BC 버튼: 12개 (1:1 문의하기 — 메타데이터 #{type}_#{template}_#{name})
- 퀵버튼: 4개 (리마인드용 1:1 문의하기)
- 메타데이터 → 해피톡 상담사가 고객명+서비스+상태 즉시 파악

### 잔여 작업
- [ ] 솔라피 검수 제출 (17종 전체)
- [ ] Make Router 13개 분기 연결
- [ ] 검수 승인 후 E2E 테스트
- [ ] 검수 안정화 후 Make→솔라피 직접 호출 전환 (FLOW_change_request.md 참조)

---

## Phase 2: 대시보드 (허브 + 카테고리) ✅ 완료

**목적:** 운영자가 한눈에 현황 파악 — 허브(3초 전체 파악) + 카테고리별 전용 대시보드 3개

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 2-1 | 허브 대시보드 | ✅ | /dashboard → 주문/상담/복원수리 3개 HubCategoryCard (핵심 수치 + 클릭 이동) |
| 2-2 | 주문 전용 대시보드 | ✅ | /orders/dashboard → 파이프라인바 + 통계 4개 + 결제완료 UrgentList |
| 2-3 | 상담 전용 대시보드 | ✅ | /consultations/dashboard → 통계 4개 + 오늘 일정 타임라인 + 미확인 UrgentList |
| 2-4 | 복원수리 전용 대시보드 | ✅ | /repairs/dashboard → 파이프라인바 + 경과일 경고 + 접수/비용안내 UrgentList 2열 |
| 2-5 | 공유 컴포넌트 | ✅ | hub-category-card, pipeline-bar, urgent-list 3개 재사용 컴포넌트 |
| 2-6 | 내비게이션 업데이트 | ✅ | NAV_ITEMS matchPrefix 기반 active 판정, href → 카테고리 대시보드 |
| 2-7 | 캐시 무효화 통합 | ✅ | hub-stats / order-dashboard-stats / consultation-dashboard-stats / repair-dashboard-stats |

### 아키텍처
```
/dashboard (허브)
  ├─ 주문 카드 → /orders/dashboard (주문 전용)
  ├─ 상담 카드 → /consultations/dashboard (상담 전용)
  └─ 복원수리 카드 → /repairs/dashboard (복원수리 전용)

각 카테고리 대시보드 → "전체 목록" → /orders, /consultations, /repairs
```

### 주요 파일
- `hooks/use-dashboard-stats.ts` — 4개 통계 훅 (useHubStats, useOrderDashboardStats, useConsultationDashboardStats, useRepairDashboardStats)
- `components/dashboard/hub-category-card.tsx` — 허브 대형 클릭 카드
- `components/dashboard/pipeline-bar.tsx` — 수평 파이프라인 시각화
- `components/dashboard/urgent-list.tsx` — 긴급 건 리스트

---

## Phase 3: 아임웹 자동 동기화 📋 미착수

**목적:** 수동 동기화 버튼 없이, 주기적으로 신규 주문 자동 유입

| # | 작업 | 상태 | 설명 |
|---|------|------|------|
| 3-1 | Vercel Cron 설정 | 📋 | 5~10분 간격 자동 폴링 |
| 3-2 | 증분 동기화 최적화 | 📋 | 마지막 sync 시점 이후 변경분만 |
| 3-3 | 동기화 상태 대시보드 표시 | 📋 | 마지막 동기화 시간, 에러 로그 |

---

## Phase 4: 알림톡 연동 (Make + Solapi) 🔧 진행중

**목적:** 주문 상태 변경 시 고객에게 자동 알림톡 발송

| # | 작업 | 상태 | 설명 |
|---|------|------|------|
| 4-1 | 일반 주문 알림 | 📋 | 아임웹 자동 알림톡 활용 (TMS 개입 불필요) |
| 4-2 | 상담 알림톡 13종 분기 | ✅ | Phase 1.8에서 구현 완료 (검수 대기 중) |
| 4-3 | 복원수리 알림톡 5종 | ✅ | Phase 7에서 구현 (as_received/cost_notice/payment_confirmed/shipped/satisfaction) |
| 4-4 | Make → 솔라피 직접 호출 전환 | 📋 | 검수 안정화 후 전환 예정 (비용 절감) |

### 알림톡 역할 분리
- **일반 주문 (가위/주변제품)**: 아임웹 알림톡 (결제완료→발송→배송완료)
- **상담 (매장/출장/톡상담)**: Solapi 전담 — 17종 템플릿 (Phase 1.8)
- **복원수리 (계좌입금)**: Solapi 전담 5종 — 접수/비용안내/입금확인/출고/만족도 (Phase 7)

---

## Phase 5: 이카운트 연동 📋 미착수

**목적:** 주문/배송 데이터를 이카운트 ERP에 자동 반영 — 재고 차감, 매출 전표 생성

| # | 작업 | 상태 | 설명 |
|---|------|------|------|
| 5-1 | 이카운트 API 조사 | 📋 | 인증 방식, 엔드포인트 확인 |
| 5-2 | 출고 전표 자동 생성 | 📋 | 송장 생성 시 이카운트 출고 연동 |
| 5-3 | 재고 실시간 동기화 | 📋 | 이카운트 재고 → TMS 표시 |
| 5-4 | 매입/매출 전표 | 📋 | 결제 확정 시 자동 전표 |

---

## Phase 6: 통합 리뷰 시스템 📋 미착수

**목적:** 오프라인 상담 + 온라인 구매 리뷰를 하나의 시스템으로 통합 (아임웹 네이티브 리뷰 미사용)

| # | 작업 | 상태 | 설명 |
|---|------|------|------|
| 6-1 | Google Sheets 리뷰 DB | 📋 | 리뷰 저장소 + GAS API |
| 6-2 | 리뷰 작성 알림톡 | 📋 | 배송완료/수리완료 후 리뷰 요청 |
| 6-3 | 리뷰 위젯 (아임웹) | 📋 | 메인/상품 페이지 리뷰 표시 |
| 6-4 | 리뷰 관리 어드민 | 📋 | 승인/답변/통계 |

> 상세: `memory/REVIEW_SYSTEM_BRIEF.md` 참조

---

## Phase 7: 복원수리 관리 🔧 진행중

**목적:** 복원수리 접수→수거→검수→비용안내→입금→수리→출고 전 과정을 TMS에서 관리

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 7-1 | Supabase 스키마 + types.ts | ✅ | repairs/repair_inspections/repair_history 3개 테이블 + RepairStatus ENUM |
| 7-2 | lib/repair/ 유틸리티 | ✅ | transitions (12상태 전이), cost-calculator, inspection-text (자동 문구), sync |
| 7-3 | API Routes | ✅ | /api/repair/ (CRUD + 검수 + 출고 + 알림 + 동기화 + 리포트) 7개 라우트 |
| 7-4 | hooks/use-repairs.ts | ✅ | React Query 훅 8개 (목록/단건/상태변경/검수/출고/취소/알림/동기화) |
| 7-5 | 목록 UI + NAV | ✅ | /repairs 페이지, 상태 탭(7그룹), 카드 목록, 경과일 표시 |
| 7-6 | 상세 UI + 검수 UI | ✅ | 2컬럼 레이아웃, 검수 체크리스트(7항목), 자동 문구, 비용 안내, 출고, 타임라인 |
| 7-7 | GAS 동기화 | ✅ | Code.gs doPost AS_CREATE → TMS sync webhook 추가, 상태변경 시 양방향 동기화 |
| 7-8 | 알림톡 5종 | ✅ | as_received/as_cost_notice/as_payment_confirmed/as_shipped/as_satisfaction |
| 7-9 | Supabase SQL 실행 | 📋 | sql/phase7_repairs.sql → Supabase SQL Editor에서 실행 |
| 7-10 | GAS Script Properties 설정 | 📋 | TMS_REPAIR_SYNC_URL, TMS_BASE_URL, CRON_SECRET 설정 |
| 7-11 | 수리내역 페이지 TMS API 연동 | ✅ | page_as_report.html → GitHub Pages + TMS API(CORS) 구조로 구현 |
| 7-12 | E2E 테스트 | 📋 | 접수→검수→비용안내→입금→수리→출고 전체 플로우 검증 |
| 7-13 | PC 마스터-디테일 레이아웃 | ✅ | 좌측 목록 + 우측 상세 패널 (lg+), 모바일은 기존 페이지 이동 유지 |

### 상태 머신 (9상태, 단순화 v2)
```
방문수거: intake(수거접수 필요) → pickup_scheduled(입고대기중) → cost_notified(진행중)
         → payment_confirmed(자동→repairing 진행중) → shipped(출고 완료) → delivered → completed

직접발송: intake(접수) → cost_notified(진행중)
         → payment_confirmed(자동→repairing 진행중) → shipped(출고 완료) → delivered → completed

레거시 호환: picked_up, inspecting, ready_to_ship (기존 데이터 전이 가능)
```

### 운영 플로우
```
고객 접수 (page_form.html)
  → GAS doPost(AS_CREATE) → Sheets 저장 + Make 알림톡 + TMS sync
  → TMS /repairs 목록에 자동 반영

방문수거: [수거접수 완료] → [입고 & 비용안내](검수+비용+알림톡) → [입금확인](자동 진행중) → 송장생성(출고 알림톡)
직접발송: [입고 & 비용안내](검수+비용+알림톡) → [입금확인](자동 진행중) → 송장생성(출고 알림톡)
```

### 최근 완료 (2026-02-25)
- [x] AS 폴더 구조 정리 — _gas/ 서브폴더 제거, consulting과 clasp 패턴 통일
- [x] deprecated 파일 삭제 (index.html 46KB, RepairReport.html 16KB)
- [x] .claspignore 생성 (page_*.html, iframe_*.html, icons/ 등 GAS 제외)

### 이전 완료 (2026-02-24)
- [x] 상태 머신 단순화 (12→9상태, 알림톡 수동발송 카드 제거)
- [x] 사이드바 통합 (비용+비용안내+액션+출고 → SidebarActionCard)
- [x] 접수정보 수량/주소 인라인 수정 + 비용 자동 재계산
- [x] 검수 폼 개선 (+ 버튼 추가 방식, 사진 촬영 UI)
- [x] 목록 날짜 표시 정리 (formatRelative 제거)
- [x] 수리내역 페이지 TMS API 연동 (page_as_report.html → GitHub Pages + TMS CORS API)

### 잔여 작업
- [ ] **주소 수정 시 다음 주소검색 API 연동** (롯데택배 송장 호환)
- [ ] **사이드바 버튼 배치 순서 조정** (수거접수 완료 → 입고&비용안내 순서)
- [ ] 사진 업로드 Supabase Storage 연동 (버킷 생성 필요)
- [ ] 수리내역서 자동 생성 (Before/After 타임라인 웹카드)
- [ ] Supabase SQL 실행 (sql/phase7_repairs.sql)
- [ ] GAS Script Properties 설정 (TMS_REPAIR_SYNC_URL 등)
- [ ] 솔라피 복원수리 템플릿 5종 등록/검수
- [ ] Make Router에 복원수리 5종 분기 추가
- [ ] 사진 마킹 (photo-marker.tsx) — html2canvas 캡처 기능

---

## 범례
- ✅ 완료
- 🔧 진행중
- 📋 미착수
- ⏸️ 보류
