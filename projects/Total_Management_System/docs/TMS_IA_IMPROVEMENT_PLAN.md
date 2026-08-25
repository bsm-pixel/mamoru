# TMS IA/UX 개선 플랜 (전수 감사 종합)
> 작성 2026-08-25 · 5개 영역 병렬 전수 감사 종합 · **PC 중점** (모바일 예외: 출장 지도/네비, 판매입력, 복원수리 수리내역서)

전 화면·공용 컴포넌트를 5축(매출흐름 / 고객접점 / 상품·재고 / 대시보드·리포트·설정 / 크로스커팅)으로 감사한 결과를 **중복 제거 + 우선순위화**한 실행 플랜.

---

## 0. 한 줄 요약
- TMS는 뼈대(마스터-디테일·공용 컴포넌트·상태맵)는 잘 잡혀 있으나, **같은 개념(상태·라벨·색·상세화면)이 여러 곳에 복제되며 드리프트**가 시작됐고, **일부 화면(판매·이벤트·시리얼·리포트·리뷰)이 표준에서 이탈**해 있으며, **크로스 도메인 이동(특히 복원수리·타임라인)** 이 끊긴 게 핵심.
- 가장 임팩트 큰 건 **"공용화 6종"** — 한 번 고치면 15~40개 화면이 동시에 정돈됨.

---

## P0 — 버그·데이터 정합 (즉시, 낮은 리스크)

| # | 항목 | 위치 | 왜 |
|---|---|---|---|
| P0-1 | **매입 상세 풀페이지 Rules-of-Hooks 위반 → 런타임 크래시** | `purchasing/[id]/page.tsx` (조기 return 후 useState) | 리포트(`reports/page.tsx:536`)에서 이 라우트로 진입 시 "rendered more hooks" 크래시. useState를 return 위로 이동 |
| P0-2 | **리포트 거래내역서 월초 날짜 UTC 버그** | `reports/transaction/page.tsx:31-32` (`toISOString().slice(0,10)`) | 과거에 고친 KST 월초 오차 재발. `toLocalDateString`으로 |
| P0-3 | **저재고 기준값 화면 불일치** | `products/page.tsx:64` 하드코딩 `<=3` vs `inventory` 설정값 | 임계치 바꿔도 제품화면 뱃지 안 바뀜. 제품도 `useSetting('inventory.low_stock_threshold')` |
| P0-4 | (모바일) 출장 일별뷰 `today` UTC off-by-one | `mobile-field-day-view.tsx:14` | 자정 전후 날짜 하이라이트 틀어짐 |

---

## P1 — 공용화 (최대 임팩트, 여러 화면 동시 개선) ⭐

### P1-1. 라벨 사전 SSOT — `lib/utils/labels.ts` 단일화
같은 라벨이 5~10곳에 복붙되어 **이미 값이 어긋남**:
- 결제수단(card/cash/transfer/mixed): 10곳 재정의 (`sales/page:30`, `sale-detail-panel:27`, `contracts:49`, `reports:14`, `api/reports/export:126` …)
- 상담유형(store_visit): `format.ts`'매장 방문' vs `customer-quick-modal:12`'매장방문' vs `dashboard-calendar:32`'매장' — **3종**
- 판매채널: `format.ts` vs `sales/[id]:22` vs `sale-detail-panel:42` — **3종**
- 결제라벨 CMS: 목록'CMS' vs 패널'CMS 자동이체'
→ **한 파일에서 export, 재정의 전부 제거.** (즉시 워딩 통일 + 15+ 파일 정리)

### P1-2. 공용 상태 배지 `<StatusBadge>` — 색 어휘 통일
- 인라인 `bg-*-100 text-*-700` 패턴 **75회/21파일** vs 공용 `Badge`+시맨틱토큰 혼재
- **딜러 색이 화면마다 다름**(판매=파랑 vs 고객=보라), 판매 1행은 **색줄+도트+텍스트색+라벨 4중 인코딩** + 6색 범례
→ `{도메인, 상태값}` → 라벨/색 자동 적용하는 `<StatusBadge>` 도입. 판매 4중인코딩을 배지 1개로 축약.

### P1-3. 상세화면 이중구현 정리 — "패널이 단일 진실, 풀페이지는 얇은 래퍼"
목록 우측 **패널**과 `/[id]` **풀페이지**가 공존하며 후자가 스테일/빈약/버그:
- 주문 `orders/[id]`(패널 대비 직접수령·주문취소 없음), 판매 `sales/[id]`(구 채널어휘 3벌, 발송·리뷰 없음), 매입 `purchasing/[id]`(부분입고·검수·인쇄 없음 + P0-1 크래시), 제품 `products/[id]`(유령, 네비 도달 불가)
→ 풀페이지를 **패널 재사용 래퍼**로 교체하거나 삭제. (진입 경로별로 다른 UI 만나는 문제 해소)

### P1-4. 팔레트 2세대 통일 (stone 계열)
- 신(stone/emerald/rose/amber): dashboard·repairs·sales·orders …
- 구(terracotta/cream/indigo-black/warm-ivory): **reports·reports/transaction·reviews·tax-invoices**
→ 회계·리뷰류를 stone 계열로. (한 사장님이 오가는 화면이 다른 브랜드처럼 보이는 문제)

### P1-5. 표 렌더 `DataGrid`로 수렴 + 死文 정리
- 공용 DataGrid: orders·purchasing·contracts·customers·suppliers·deliveries
- 자체 표: **sales(`SalesGridTable`)·inventory(`InventoryTable`)** → DataGrid 이관(원래 의도)
- **`DataTable`(components/ui/data-table.tsx) = 완전 사문** → 삭제
- products 카드그리드는 정책상 유지 가능

### P1-6. 요약카드 `StatCard` 단일화 + 금액/포맷 통일
- 요약카드 3종(StatCard / RevenueDarkCard / 생 Card 수기): tax-invoices·reports 등의 생 Card → StatCard
- **대시보드 금액 만원반올림**(`₩12만`) vs 나머지 `formatKRW` 원단위 → 통일
- `toLocaleString` 직접 20회(serials·events·purchasing/new·reviews) → `formatKRW`

---

## P2 — IA 연동 (CRM 핵심 동선) ⭐

### P2-1. 크로스 도메인 이동 복구
- **고객 상세 타임라인에서 상담·복원수리 클릭 불가** (`customer-detail-panel.tsx:502` — sale/contract/order/invoice만 라우팅). → consultation→`/consultations/[id]`, repair→`/repairs/[id]` 추가 **(HIGH, 고객 360뷰 핵심)**
- **복원수리가 연동 고아**: 상담/판매 상세엔 복원수리 링크 없음, 복원수리 상세엔 상담/판매/고객 링크 없음
- **전환 역링크 부재**: 이벤트→그 판매, 계약→그 판매(`offline_sale_id` 있는데 링크 없음), 온라인주문↔판매
- (방금 한 주문↔고객 링크·활동칩의 연장선 — 같은 원리로 상호 딥링크)

### P2-2. 대시보드 IA 재정비
- **드릴다운 오연결**: 저재고→`/purchasing/new`(발주폼) → **재고 목록**으로 / 운송장→`/settings` → **발송 화면**으로
- **설정↔대시보드 미반영**: 카드배치 설정(outstanding/lowStock/todos)이 실제 렌더에 안 먹음 + 기본순서 불일치
- **상단 IA**: 월간 캘린더가 화면 중앙 점유 → 핵심 "미수금·긴급·할일"이 스크롤 아래로. 핵심 액션을 상단, 캘린더 축소/우측
- 매출 KPI 도넛이 RevenueDarkCard 규칙 위반 → 다크카드 계열로
- 에러 상태 부재(장애 시 전부 "0"으로 보여 정상 오인) → isError 분기
- 일정 페이지와 대시보드 캘린더 100% 중복 → 역할 분리(대시보드=요약 스트립, 캘린더=일정 전용)

### P2-3. 아임웹/ALPS 왕복을 "지금 할 일 1개"로
- "아임웹 배송대기 처리 먼저" 토스트, "1시간마다 자동확인" 반복 안내, ALPS 취소확인 수동버튼 등 시스템 사정이 날것 노출
- 판매 상세 발송 섹션이 상태별 버튼 최대 6개 수직 적층 → **스텝형(현재 단계 1 + 다음 액션 1)**
- 재고조정/입고 시 아임웹 재고 자동동기가 UI에 흔적 없음 → "아임웹에도 ±N 반영" 안내 + 연동품목 뱃지

---

## P3 — UI 표준 수렴 (PC 활용)

| # | 항목 | 위치 |
|---|---|---|
| P3-1 | **이벤트·재고판매·시리얼 PC 마스터-디테일 없음** (PC에서도 SlidePanel/1열) → 좌목록+우패널 도입 | events·stock-sale·`serials/page.tsx:150` |
| P3-2 | **톡상담 탭만 우측 상세패널 없음** (형제 탭과 유일 이탈) | `consultations/page.tsx:306` |
| P3-3 | **전체상담 탭 저밀도 카드** vs 고객·복원수리 표 → 밀도 통일 | `all-consultations-list.tsx` |
| P3-4 | 세금계산서 → 발행목록 좌 + 상세 우 (좌우 적합) | `tax-invoices` |
| P3-5 | `useIsLg` 훅으로 인라인 `matchMedia` 4곳 일원화 | sales·inventory·consultations·products |
| P3-6 | 우패널 폭 통일(440px 표준; consultations 400·inventory 55%·sales 가변) | 각 상세 |

---

## P4 — 정리·폴리시 (낮은 리스크, 유지보수)

- **데드코드 제거**: `repairs/tabs/*` 6개 + `repair-action-chips`·`inspection-photo-marker`·`repair-photos`·`repair-tab-bar`(미사용), `DataTable`, `ContractTableRow`, 고아 라우트(`orders/dashboard` 네비 미노출, `repairs/dashboard`·`consultations/dashboard` redirect-only, `products/[id]` 유령)
- **파괴적 액션 다이얼로그 통일**: `window.confirm/alert`(events·tax-invoices·manual-invoices·sale-detail 일부) → `ConfirmModal`
- **로딩/빈 통일**: 상세 "불러오는 중…" 텍스트 → Skeleton, `EmptyState` 전면 적용
- **네비 그룹핑 IA**: 복원수리가 '상담' 그룹에(→ CS 성격), 빠른송장이 '판매'에(→ 배송 성격), 계약서/세금계산서 분리 검토. B2B 명칭 3종("B2B거래처"/"B2B거래"/"거래처") → 1개 확정
- **제목 규약**: 사이드바 라벨 = Topbar title 일치. 주 생성버튼은 Topbar `action` 슬롯으로
- **sidebar/mobile-nav 아이콘맵 이중** → 공유
- **끊긴 안내 링크**: `sales/page.tsx:544` "[B2B거래] 메뉴에서" (삭제된 메뉴 지칭) → `/deliveries` 직접
- **모바일 예외 정리**: 출장 네비 이원화(웹검색 vs `kakaomap://route`) → 좌표기반 route로 통일 / 지도 InfoWindow에 네비 버튼 + 빈상태 안내 / 핀(점) 삭제 오탭(투명 32px) → 삭제 어포던스
- **의미 라벨**: 시리얼 "창고 이동"이 실은 생성/삭제 겸함 → "재고→시리얼 전환" 뉘앙스 / "입고&비용안내" 단일버튼에 입고 단계 흐림

---

## 권장 실행 순서 (라운드)
1. **R0**: P0 (버그 4건) — 반나절, 리스크 0
2. **R1**: P1-1(라벨 SSOT) + P1-2(StatusBadge) — 최대 파급, 이후 작업의 토대
3. **R2**: P1-3(상세 패널 단일화) + P4 데드코드 — 구조 정리
4. **R3**: P2-1(크로스 링크) + P2-2(대시보드 IA)
5. **R4**: P1-4(팔레트) + P1-5(DataGrid 수렴) + P3(PC 표준)
6. **R5**: P2-3(발송 스텝화) + P4 나머지 폴리시

각 라운드는 독립 배포 가능. R1·R2가 "손대는 파일 수 대비 효과"가 가장 큼.

---

## 🟢 진행 로그 (2026-08-25 실행)

| 커밋 | 내용 | 상태 |
|---|---|---|
| `9aa2d60e` | **R0** P0 버그 4건(훅 위반 크래시·거래내역서 UTC·저재고 기준·출장 UTC) | ✅ 완료·검증 |
| `6143371f` | **R1(P1-1)** 라벨 SSOT — 고객유형·상담유형 색/라벨 단일화(딜러색 충돌·상담 3종 드리프트 해소, 7파일) | ✅ 완료·검증 |
| `4766d459` | **R2(P2-1/P2-2)** 크로스링크 — 타임라인 전 이력 클릭·계약→판매·이벤트→판매 딥링크·저재고 드릴다운 교정 | ✅ 완료·검증 |
| `195eb430` | **P4** 데드코드 873줄 제거(repairs/tabs 6·repair-action-chips·inspection-photo-marker·data-table·ContractTableRow) + 끊긴 안내링크 교정 | ✅ 완료·검증 |
| `8499cbd1` | **P1-4** 팔레트 통일 — 회계·리뷰류 구팔레트→stone 38곳 | ✅ 완료·검증 |
| `398fac93` | **P4** 네비 아이콘맵 SSOT — sidebar·mobile-nav 이중정의 통합(nav-icons.ts) | ✅ 완료·검증 |
| `0dc277b4` | **P1-3** 상세 패널 단일화 — sales/[id]·purchasing/[id] → 패널 래퍼, products/[id] 유령라우트 리다이렉트(839줄↓) | ✅ 완료·검증(orders는 보류) |
| `c66e015a` | **P3-2** 톡상담 탭 PC 2열(리스트\|상세) — 형제 탭과 통일 | ✅ 완료·검증 |
| `c1e6871f` | **P3-1** 시리얼·이벤트 PC 마스터-디테일 2열 + B2B 명칭 통일(납품↔거래처) | ✅ 완료·검증 |
| `c77c18bd` | **P2-2** 대시보드 설정↔렌더 정합(유령 카드 제거·기본순서 일치·저장값 자가치유) + 장애 경고 배너 | ✅ 완료·검증 |
| `106132cc` | **P3-5** 인라인 matchMedia → useIsLg 훅 일원화(consultations·inventory·products) | ✅ 완료·검증 |

각 커밋 `tsc 0 / build 성공` 확인.

**P1-3 세부**: sales·purchasing은 패널이 구 풀페이지의 상위집합임을 grep 검증 후 스왑(안전). products/[id]는 도달 경로 0 유령라우트라 리다이렉트. **orders/[id]는 보류** — 패널이 OrderActionBar를 포함해 스왑 자체는 가능하나, 구 페이지의 `cancel_pending 마운트 시 자동 ALPS 취소확인` 편의가 패널엔 없고 ALPS 플로우 실측 불가 → 액션 패리티 확인 후 별도 진행.

### 판단으로 보류(사유 명시)
- **P1-2 StatusBadge 전면 sweep**: 실제 결함(딜러 색 충돌)은 P1-1에서 SSOT로 해소됨. 판매행 4중인코딩(색줄+도트)은 **2026-05-26 사장님 채택 디자인**이라 Fix Preservation 원칙상 임의 개편 안 함.
- **P4 모바일 출장 네비 통일**: PC=웹지도 링크 / 모바일=`kakaomap://route` 앱네비로 컨텍스트가 달라 일괄 통일 시 PC에서 앱URI가 죽음. 사장님 실사용 플로우 + 헤드리스로 앱 동작 실측 불가 → **디바이스 확인 후 진행 권장**.
- **P1-6 금액 포맷 통일**: 외국통화(¥·환율)가 다수 섞여 있고 `₩`접두 vs `원`접미 혼재라 일괄 치환 시 표기 불균일 위험 → 정밀 작업 필요.

### 남은 라운드 — ⚠️ 시각/구동 검증이 커서 사장님 확인 하에 진행 권장 (TMS는 인증 게이트라 헤드리스 캡처 불가 → blind 변경 시 시각 회귀를 못 잡음)
- **P1-5 DataGrid 수렴**: sales `SalesGridTable`·inventory `InventoryTable` → DataGrid. ⚠️ **실사용 대형 목록 리렌더 → 시각 QA 필수**. (sales는 색줄/도트 '안 A' = 사장님 채택 디자인이라 보존 우선)
- **P2-3 발송 스텝화**: 판매 상세 발송 버튼 6개 적층 → 스텝형. ⚠️ **2026-05-26 사장님 채택 판매 디자인 영역 + 발송 플로우 실측 불가 → 합의·시연 후**
- **P4 window.confirm/alert(20+) → ConfirmModal/toast**: ⚠️ **파괴적 액션(삭제·취소) 플로우라 건별 실측 필요**
- **orders/[id] 단일화**: OrderActionBar 액션 패리티(cancel_pending 자동 ALPS 체크) 확인 후
- **소품목**: P1-6 금액 포맷(외국통화 혼재 주의), P3-3 전체상담 밀도, P3-4 세금계산서 좌우, P3-6 우패널 폭(400/440) 통일, 네비 그룹핑, 매출 KPI 도넛 다크카드화

> **원칙**: 이번 세션은 "검증된 기존 패턴 복제 + 로직 SSOT/버그픽스"만 자율 진행(헤드리스로 검증 가능한 범위). 위 남은 항목은 **신규 시각 디자인 발명·승인 디자인 개편·파괴적 플로우 재배선**이라 사장님이 화면을 보며 확인하는 게 사고 예방에 맞음.
