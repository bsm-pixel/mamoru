# MAMORU 시스템 구축 — TODO

## 📅 할일 빠른 확인 (날짜별 · 미완료만)

> 상세는 아래 각 섹션 참고. **⚠️=사장님 외부작업 / 💻=클로드 작업 / ⭐=검토 우선 / (선택)=급하지 않음**

### 2026-08-30 · 주문 제품 교환 (배포완료 → 후속)
- [ ] 💻 **반품 → 재입고(판매가능) 경로** — 반품창고 복원 가위를 판매재고로 올리는 버튼
- [ ] 💻 **판매관리 교환 다품목 확장(B)** — 주문교환 모달을 공통 picker로 추출, offline_sales에 1→N 적용
- [ ] 💻 (선택) 死코드 정리 — `useExchangeShip` 훅 + `/api/orders/[id]/exchange-ship` 라우트(Phase2에서 UI 제거됨, 교환 송장 SSOT=returns)
- [x] ✅ 교환 흐름 발송 ✓ = '집하' 연동 (23e464b1 배포·전미하 데이터 정정)

### 2026-08-03 · 리뷰 이벤트 (배포완료 → 세팅 대기)
- [ ] ⚠️ 아임웹 '리뷰 이벤트' 페이지 생성 + 코드위젯 붙여넣기
- [ ] ⚠️ TMS `/reviews/event` → 상품·마감 입력 → [게시] → 월말 [당첨 발표]
- [ ] (선택) 당첨 알림톡 연결

### 2026-07-24 · 복원수리 알림톡 3건 (점검)
- [ ] ⚠️ 매장방문 예약완료 알림톡 확인/신설(`as_visit_booked`)
- [ ] 💻 방문 전 리마인드 알림톡(크론+템플릿)
- [ ] ⚠️ 어드민 수동 접수 시 알림톡 발송 여부 확인

### 2026-07-12 · 집하 자동감지 (관찰중)
- [ ] ⚠️ 감지 타이밍 검증(ALPS 대조) → 검증되면 출고 알림톡 ON
- [ ] 💻 ALPS 집하 시각 필드명 확정 · 크론 실행시간 관찰

### 2026-07-12 · EVENT폼·회사소개·빌더 잔여
- [ ] ⚠️ EVENT 접수폼 코드위젯 붙여넣기 + 실제 접수 1건 테스트
- [ ] ⚠️ 회사소개 CTA 목적지·지도이미지·모바일 퀵네비 실기 확인
- [ ] 💻 빌더 DRY 가격로직 · 틴닝 SVG/카피

### 2026-07-11 · 회사소개 자산
- [ ] ⚠️ 기능 아이콘 6개(arrow/kakao/instagram/map/nav/pin) 파일 배치
- [ ] ⚠️ 히어로 로고 결정(현재 빈 src)

### 2026-07-10 · 카테고리 페이지 위젯 전환
- [ ] ⚠️ 페이지별 **d01+n10 추가 → 확인 → 기존 코드위젯 삭제**(순서 엄수)

### 2026-07-08 · 메인페이지 커스텀 위젯 전환
- [ ] ⚠️ 위젯 아임웹 등록·테스트 → 피드백 · 배경 리듬 세팅 · iframe 중복 제거

### 2026-05-31 · 샘플소싱 (선택)
- [ ] 💻 Phase4 파일 공용 이동 · ⚠️ 라벨 프린터 기종 최종 선정

### 2026-05-27 · TMS 톤 통일 (선택 잔여)
- [ ] 💻 배경 `bg-stone-50` 미적용분 · 설정 카드 라운드 미세조정

### 2026-05-25 · OS-004 후기요청 진단
- [ ] ⚠️ 솔라피/Make `purchase_review_request` 3항목 확인

### 진행/검수 대기 (날짜 유동)
- [ ] ⭐ 상품 카탈로그(`products/catalog/preview.html`) 검토
- [ ] ⚠️ 복원수리 직접방문 Phase4 — 솔라피 알림톡 5종 검수

---

> 최종 수정: 2026-08-25 — **TMS IA/UX 전수감사 개선 R0~P1-4 실행**(5커밋, 상세 `projects/Total_Management_System/docs/TMS_IA_IMPROVEMENT_PLAN.md`). ①R0 버그4건(훅크래시·UTC날짜·저재고기준) ②R1 라벨SSOT(고객/상담유형 색·라벨 단일화, 딜러색 충돌 해소) ③R2 크로스링크(고객 타임라인 전이력 클릭·계약/이벤트→판매 딥링크·저재고 드릴다운) ④P4 데드코드 873줄 삭제 ⑤P1-4 팔레트 stone 통일. 각 tsc0/build성공. **남은 구조 대공사(P1-3 패널단일화·P1-5 DataGrid수렴·P3 PC마스터-디테일·P2-3 발송스텝화)는 시각/구동 검증 커서 별도 회차 권장** — 플랜 문서 진행로그 참조.
>
> 이전: 2026-07-16 — **TMS PC 그리드(밀집표) 토글 전면 확산**. 판매관리 검증 패턴을 공용화(`useGridMode` 훅 + `components/ui/data-grid.tsx` DataGrid/GridToggleButton)하여 **복원수리·고객·매입·주문·B2B납품·계약·거래처 7개 화면**에 이식(카드뷰·모바일 유지, 그리드모드 시 목록 넓게/상세 420px 반전). 각 화면 인라인 matchMedia effect도 훅으로 통합. 판매관리 그리드 컬럼 재배치(날짜·고객·상태·상담유형·배송상태·금액)+리뷰약속 세로깨짐 컨테이너쿼리 fix. consultations는 지도/달력 특수레이아웃이라 제외.
>
> 이전: 2026-07-12 — **EVENT 접수폼 아임웹 iframe 삽입**(임베드 모드로 고정CTA·모달 해결) · **회사소개 연락처/링크 fix** · **상품 빌더 서체 규칙(영문숫자 Paperlogy / 한글 Pretendard)** · **틴닝 스펙 PC 4열 재설계** · **DRY 카드 일원화**.
>
> 이전: 2026-07-10 — **카테고리 페이지 위젯화**(d01 배너 겸용 + n10 현재페이지 칩) · **scroll-snap이 모바일 좌측여백 먹던 근본원인 해결**(n10·n12·t02) · **n15 상품탭 폐기**(Shadow DOM 격리) · **회사소개 외부이미지 2장 저장소 이관 + 대표사진 92% 경량화** · **메인 후기 자기완결 iframe 분리**.
>
> 이전: 2026-05-27 — TMS 톤 통일 4그룹 + 공통 컴포넌트 추출(StatCard/RevenueDarkCard) + 리뷰 약속 유형/subtype(094/095) + IA 모순 해결 + 알림톡 진단 사전 안내 룰 박제.
>
> 완료 이력은 모두 git history(`git log --oneline`) 참조. 본 파일은 진행 중·대기만 유지.

---

## 🟡 주문 제품 교환 (2026-08-30) — 기능 배포 완료, 후속 대기

**완료·배포**: 아임웹 주문 **[제품 교환]**(매출·카드 불변, 상품/재고만 스왑) + 다품목 + 발송방식(배송/직접전달) + **[교환품 송장 생성]**. 마이그 140·141 실행됨. 재고 드리프트 3버그 수정(**현재고 명시차감**·반품 **멱등**·배정목록 필터) — 전미하 실사고로 검증·데이터 정정 완료. 기억 `reference_order_exchange`, 매뉴얼 `MANUAL_ORDER §11`.

**후속(대기)**:
- [ ] **반품 → 재입고(판매가능) 경로** — 반품창고의 복원된 가위를 다시 판매재고로 올리는 버튼(있는지 확인 후 없으면 제작)
- [ ] **판매관리 교환 다품목 확장 (B)** — 주문교환 모달을 공통 picker 컴포넌트로 추출해 offline_sales 교환에도 적용(현재 1→1 한계를 1→N으로). 사장님 승인받음, 미착수.

## 🟢 리뷰 이벤트 — 구축 완료, 사장님 세팅 대기 (2026-08-03)

고객 페이지(`page.mamoru.kr/projects/reviews/page_review_event.html`) + TMS 「리뷰 이벤트 관리」(`/reviews/event`) 완전 자동 연동(fetch) **배포 완료**(91535fff). 지난 당첨자=월 탭 아카이브. 마이그 123 실행됨.

**⏳ 다음(사장님 작업)**:
1. **아임웹**: 이벤트 하위 '리뷰 이벤트' 페이지 생성 → 코드위젯에 `projects/reviews/iframe_review_event.html` 붙여넣기 → N13 이벤트목록 카드 링크를 그 페이지로.
2. **TMS**: `/reviews/event` → 8월 상품(1·2·3등)·마감일 입력 → **[진행중으로 게시]** (→ 고객 Hero·이달의 상품 노출). 월말 후기 심사 → 등수 선정 → **[당첨 발표]** (→ 지난 당첨자 노출).
3. 첫 저장/발표 시 실로그인 상태 정상동작 최종 확인(관리 POST는 헤드리스 E2E 미검증분).

**후속(선택)**: 당첨 알림톡(발표 시 당첨자에게) 미연결 · 인스타 응모(Phase 2) 보류.

---

## 🟡 진행중 — 쇼핑 카테고리 페이지 코드위젯 → 커스텀 위젯 (2026-07-10)

**목표**: 카테고리 페이지(블런트 `/29` · 틴닝 `/30` · 장가위 `/39` · 슬라이싱 `/40` · 전체보기 `/44`) 상단 코드위젯(배너+칩+fadeJS)을 **d01(배너) + n10(칩)** 으로 교체.

**⏳ 다음(사장님 작업)**: 페이지마다 **d01 추가 → n10 추가 → 확인 → 코드위젯 삭제** (순서 엄수, 먼저 지우면 배너 소실)
- d01 설정: `세로 위치=아래` · `정렬=왼쪽` · `배경 느린 줌=끄기` · `모서리=0px` · `최대 가로폭=1200` · 높이 비움
- n10 설정: `테마=라이트` · `정렬=가운데` · 칩 5개 동일 + **그 페이지 칩만 `현재 페이지` ON**
- ⚠️ 기존 배치 위젯엔 새 필드가 안 뜰 수 있음(패널 유형 고착) → **위젯 삭제 후 재추가**
- ⚠️ 아임웹 재붙여넣기 필요: **n10(HTML+CSS) · n12(CSS) · t02(CSS) · n13(CSS) · d01(HTML+CSS)**

**❌ 폐기 — `n15_product_tabs`** (저장소엔 잔존): 칩 클릭 시 상품진열 in-page 필터/전환을 시도했으나 **커스텀 위젯은 Shadow DOM 격리 + 아임웹이 상품 연동·위젯간 통신 금지** → 위젯 JS가 바깥 `#container_<WID>`에 접근 불가. **연동은 카테고리 페이지 URL 링크가 유일한 방법.**

**🔧 근본 수정 기록**: 모바일 좌측 여백이 계속 사라지던 원인 = **`scroll-snap`이 첫 항목을 컨테이너 왼쪽(0px)에 스냅해 트랙 패딩을 먹음**. `scroll-padding-left:20px` 로 해결(n10·n12·t02 전수 적용). n15는 snap 미사용.

---

## 🔵 관찰중 — 집하 자동감지 (2026-07-12 배포, 마이그 109)

배포 완료 (`87f209ac` + `4200c171`). 크론 1시간마다 집하 감지 → 자동 출고완료. **출고 알림톡은 OFF로 시드됨**.

- **🔲 1~2일 뒤 감지 타이밍 검증** — 아래 SQL로 뜨는 건들의 `shipped_at`이 실제 기사님 수거 시각과 맞는지 ALPS 화면과 대조
  ```sql
  SELECT sale_number, customer_name, customer_type, shipped_at, shipped_source
  FROM offline_sales WHERE shipped_source='alps_pickup' ORDER BY shipped_at DESC LIMIT 20;
  ```
- **🔲 검증되면 알림톡 ON** — 설정 → 알림 → [판매 출고 안내] 토글. 그때부터 B2C 고객에게 출고 알림톡 자동 발송
- **🔲 ALPS 집하 시각 필드명 확정** — 문서에 없어 후보 키 순차 탐색 중(못 찾으면 감지 시각으로 대체, 기능 영향 X).
  `curl -H "Authorization: Bearer $CRON_SECRET" ".../api/cron/track-delivery?debug=1"` → `salesPickupFirstResults[].trackingKeys` 확인 후 `alps-client.ts SCAN_DATE_KEYS` 확정 + 진단 필드 제거
- **🔲 크론 실행시간 관찰** — 블록 6개 × 50건 순차 fetch. 타임아웃 나면 신규 블록만 병렬화

---

## 🟠 별도 체크 — 복원수리 알림톡 3건 (2026-07-24, 사장님 요청)

목록 개편·직접수령·리뷰 delivery 칩은 배포 완료. 아래는 **알림톡 존재 여부 확인 후 필요시 구축** — 검수 1~3영업일 소요라 미리 점검.

- **🔲 직접방문(매장방문) 예약완료 → 접수완료 알림톡 (`as_visit_booked` 신설?)** — 현재 매장방문 예약 처리 시 고객에게 접수/예약완료 알림톡이 **가는지 불명**. [reference_solapi_templates] 27~31(당일수리 5종)에 예약완료가 포함됐는지 대조 → 없으면 신설.
- **🔲 방문 전 리마인드 알림톡** — 예약 전날/당일(24h·2h 등) 리마인드. 컨설팅 리마인드 패턴 재사용 가능한지 검토. 크론 트리거 + 템플릿 신설 필요.
- **🔲 어드민 수동 접수 시 접수완료 알림톡 미발송 여부** — 아임웹 통합접수(/53)는 `as_received` 발송되나, TMS에서 사장님이 **직접 수기 접수**한 건도 발송되는지 확인. 안 되면 접수 저장 경로에 발송 추가.
- 참고: 자동 출고완료 알림톡(집하 스캔 → `as_shipped`)은 **이미 동작 확인됨**(크론 track-delivery 블록 2-A). 복원수리도 판매와 동일하게 자동 발송 중.

---

## 🟡 대기 — 2026-07-12 작업의 남은 것

**EVENT 접수폼 (타사가위 팡팡)**
- **🔲 사장님: 코드위젯 붙여넣기** — `projects/event/iframe_event_pangpang.html` 전체 → 아임웹 이벤트 페이지 코드위젯. 위젯 속성 **'그리드 사용 안 함(전체 너비)'**. 이벤트 바뀌면 파일 복사해 `iframe_event_<이벤트명>.html` 로 새로 만든다.
- **🔲 실제 아임웹에서 접수 1건 눌러보기** — 임베드 모드(고정 CTA·완료화면·주소검색 앵커)는 로컬까지만 검증함. 캠페인 fetch·다음 우편번호는 라이브 확인 필요.

**회사소개 (`page_intro.html`)**
- **🔲 CTA "편하게 문의하기" 목적지 확정** — 현재 카카오 채널(`_KHWNb`) 새 창. 기존엔 아임웹 `/52`(상담 문의)였음. `/52`가 맞으면 한 줄 되돌리면 됨.
- **🔲 지도 이미지** `projects/brand/img/location_map.jpg` 없음 → 안내문 표시 중. 파일 넣으면 자동 표시. (`img/2222.jpg`가 그 지도인지 미확인 — 미추적 상태로 방치 중)
- **🔲 모바일 퀵네비 칩 실기기 확인** — iframe 안 앵커가 죽어서 `scrollIntoView`(상위 프레임 전파)로 배선함. 아임웹에서 실제로 눌러봐야 확정.

**상품 상세 빌더**
- **🔲 DRY 옵션 가격 계산 로직 미구현** — `blade_edge_dry`에 가격만 데이터로 보관(ST 0 / SD +1만 / SO +2만). 상품 페이지에 "+10,000원 옵션" 표기하려면 별도 구현 필요.
- **🔲 틴닝 카드 SVG·문구** — 홈 형태 SVG, 카피 풀(틴닝·장가위·드라이) 여전히 미작성. 히어로 `H1/T28` 코드 어색 건도 미해결.
- ⚠️ **카탈로그 `id`는 식별자다** — 편집기에서 `"24발"` → `"24"`/`"26 발"` 처럼 바꾸면 spec 매칭이 깨진다. 카드에는 어차피 숫자만 표시되므로 id는 단위 포함으로 유지할 것.

---

## 🟡 대기 — 회사소개(`projects/brand/page_intro.html`)

- **🔲 아이콘 파일 6개 미배치**(2026-07-11): 인라인 SVG → `./icons/` 파일 참조 전환 완료. 콘텐츠 아이콘은 사장님이 채움. **CTA·오시는길 기능 아이콘 6개 파일 대기**: `arrow.svg`(문의 화살표)·`kakao.svg`·`instagram.svg`·`map.svg`·`nav.svg`·`pin.svg`. 파일 없으면 로드실패 자동숨김(JS) 상태 → **폴더에 넣으면 자동 표시**(HTML 무수정). ⚠️ img로 부른 SVG는 색 상속 X → 파일에 색 구워야 함(map/nav는 hover 시 검정배경이라 중간회색 권장).
- **🔲 blunt.svg**: 기술력 카드에서 예전 참조 흔적(현재 face1_2/phone/gogo로 교체됨 — 불필요).


- **히어로 로고가 비어 있음** (`691줄 <img src="">`) — 빈 src는 깨진 아이콘 유발. 보유 자산 `img/logo.png`(656×664, 투명배경)를 72px로 리사이즈해 넣을지, 별도 화이트 로고를 쓸지 **사장님 결정 대기**.
- `img/logo.png` 656×664 (권장 200×200) — 용량 19KB라 급하지 않음. 최적화 미실시.
- 작업트리에 `img/2222.jpg`(원본 229KB) 미커밋 잔존 — `repair_work_800.jpg`로 대체됨. 불필요 시 삭제.
- **iframe 유지 결정**: 1,083줄·섹션 10개·스크롤 연출(IntersectionObserver 2·keyframes 5)이 단일 문서에 묶여 있어 위젯 분할 시 연출 파편화. fetch 0이라 기술적으론 가능하나 ROI 낮음. **SEO(부모 페이지 색인)가 중요해질 때만 재검토.**

---

## 🟡 진행중 — 메인페이지 커스텀 위젯 전환 (2026-07-08)

**목표**: 메인 top/btm iframe 임베드 → 아임웹 네이티브 커스텀 위젯으로 교체. 위젯은 `projects/imweb_widgets/`, 메인용 순서 스냅샷=`_main_page/`(번호폴더 01~16 + `_00_조립순서.md`).

**오늘 표준화 완료(커밋 be7d397d~4d6d6f0a)**:
- 인라인 style `{{}}` 전면 제거(data-*+JS 표준), 라운드 각지게 고정
- **가로 최대폭 = 자유 숫자입력 + `applyMaxw`+`MutationObserver`(편집기 실시간, 영역확장해도 설정폭)** — d01·n09·n10·t11·n12·n11 적용
- **PC/모바일 각각**: d01(이미지 비율 자동 높이 PC/모바일 독립), n09(가로폭·한글두께), n10(중앙정렬·모바일 가로스크롤peek·PC/모바일 폰트pt·상하여백·**아이콘 제거**=아임웹 SVG업로드 불가), n12(1행 가로진열·PC/모바일 이미지·높이 각각), n11(PC 열수·모바일 가로스크롤peek), t11(롤링 로고 띠 신규·끊김 해결)

**⏳ 다음(사장님 확인 대기)**:
1. 사무실/사장님이 위젯들 아임웹 등록·테스트 → 피드백 반영. **필드 구조 바뀐 위젯(n10/n12/n11)은 아임웹서 위젯 삭제 후 재추가**해야 새 필드 깔끔 생성(패널 유형 고착).
2. **MutationObserver 편집기 실시간 반영 여부 실측 확인** → 되면 "패널 스타일값=data-*+JS+Observer" 표준으로 메모리 승격.
3. 배경 리듬 세팅표: Cream 기본 + **오프닝(d01·n09)·클로징(d04·n08)만 Void 다크** + 중간 Shell/Parchment 명도차 (Brand Guide B-01). 위젯별 theme 값 확정 대기.
4. 메인 iframe→위젯 교체 시 **중복 제거**(옛 코드위젯 A/B 삭제) + 위젯 폭 통일. ⚠️ `common_code/header_code.txt`(상품진열·기획전·타페이지 공유)는 **절대 안 건드림** [feedback_imweb_common_code_scope].
5. 나머지 위젯도 필요 시 가로최대폭·PC/모바일 패턴 적용(사장님 지목).

---

## 💻 노트북에서 이어가는 방법 (사장님 PC 이동 시)

1. **git pull** — 모든 코드·docs·.claude/ ADDENDUM 자동 동기화
2. **claude code 실행** — CLAUDE.md 자동 로드 → 키워드 트리거로 ADDENDUM_IMWEB/ADMIN 자동 로드
3. **현재 컨텍스트 파악**: 이 파일 "🟡 진행중" 섹션 읽기 → 즉시 어디서 멈췄는지 파악
4. **메모리 복원 (선택, 새 PC)**: 사장님이 "MEMORY.md 인덱스 + 핵심 메모리(시리얼 무결성·iframe 패턴) 다시 정리" 한 마디면 클로드가 git 추적 파일들에서 복원

**git 추적 (노트북 자동 동기화)**:
- `.claude/CLAUDE.md` / `.claude/ADDENDUM_*.md` — 작업 가이드
- `projects/Total_Management_System/docs/TMS_FLOW_*.md` / `MANUAL_*.md` — TMS 흐름·매뉴얼
- `TODO.md` (이 파일) — 진행 상황
- 모든 `projects/` 코드

**메모리 (PC 로컬, git 추적 X)**:
- `~/.claude/projects/c--*/memory/MEMORY.md` 인덱스 + 개별 `.md` 파일들

---

## 🟡 진행중

### ✅ 샘플 소싱 선별 도구 `/sourcing` — 운영 가동 (2026-05-31 구축, Phase 1~3 + 추가기능 배포완료)

**상태**: TMS `/sourcing` 운영 가동 중. DB(098 sourcing_pos/items, 099 supplier) 실행완료. 매입→입고매칭(폰카메라/PC바코드스캐너+사진 Storage)→실테스트→선별(채택/탈락)→선별리스트출력. products INSERT·SKU·수량 없음(선별기). 소싱=회차, 품목별 회사명·회사링크·복제버튼·업체별채택현황, 라벨=QR+번호+품목명+한화가격(환율 상단)+품목별 매수.
**라벨 인쇄**: KM-106D(송장감열)로 임시 운영 가능(드라이버 용지크기 등록+갭캘리브 필수 — 메모리 교훈). 제브라 ZD421T는 ZPL 업그레이드용.

**남은 일 (선택, 급하지 않음)**:
1. **Phase 4 정리**: `LabelPreview.tsx`·`types.ts`를 design-lab/_sections → `components/sourcing/` 공용 이동 + design-lab § 1688 데모 삭제 (기능영향 0)
2. **프린터 기종 최종 선정**: ZD421T 1688직구 ~61~70만(A/S✗) vs 국내정품 77~79만(A/S○) vs TSC/Godex 국내정품 15~22만(QR엔 충분). 화면 권장문구 현재 ZD421T → 확정 시 반영
3. (확장 시) 업체 정규화 — 현재 free-text+복제로 충분

**상세**: memory `project_1688_label_printer.md`

---

### 1. 자동 후기요청 발송 진단 — `purchase_review_request` 템플릿 확인 ⚠️ 사장님 외부 작업

**배경**: OS-20260525-004 자동 cron(track-delivery) 후기요청 미발송. TMS는 success 기록(DB review_requested_at set) 했으나 솔라피에 도달 X. 사장님 수동 발송([상담→톡상담], review_request 템플릿)은 정상 도달 확인됨 — TMS→Make→솔라피 흐름은 정상.

**Make 시나리오 캡처 단서**: `제품구매_만족도후기 21` 분기가 `🚫 The bundle did not pass through the filter`. 즉 자동 cron이 보낸 `purchase_review_request` 이벤트가 Make 라우터 필터에서 차단.

**사장님이 확인할 항목 (3가지)**:
1. **솔라피 콘솔**: `purchase_review_request` 템플릿 존재 + 활성 + 검수 통과 상태인지
   - 메모리 카탈로그([reference_solapi_templates](../.claude/projects/c--Users-user-Desktop-mamoru/memory/reference_solapi_templates.md))에 2026-05-23 검수 통과 박제되어 있지만 박제 후 갱신/삭제 가능성 있음
2. **Make 시나리오** `제품구매_만족도후기 21` 모듈: 필터 조건 (event 이름·키 매핑)
3. **Make 시나리오 실행 이력 (5/27 17:00 KST ±5분)**: 자동 cron이 실제 Make 웹훅을 호출했는지

**결과별 액션**:
- 솔라피 템플릿 없음/비활성 → 사장님 신규 등록 + 검수 신청 (1~3 영업일)
- Make 필터 조건 미스매치 → 사장님 모듈 수정
- 둘 다 정상 → 클로드가 TMS payload + Make 매핑 더 깊이 진단

**관련 코드 (TMS측)**:
- [api/cron/track-delivery/route.ts](projects/Total_Management_System/app/src/app/api/cron/track-delivery/route.ts) — offline_sales 자동 발송 분기
- [lib/notification/review-request.ts](projects/Total_Management_System/app/src/lib/notification/review-request.ts) — 템플릿 매핑
- [lib/notification/make-webhook.ts](projects/Total_Management_System/app/src/lib/notification/make-webhook.ts) — Make 이벤트 페이로드

---

### 2. 🎨 TMS 톤 통일 — 시안 B+ 전체 확장 (점진 작업)

**디자인 기준**: [feedback_tms_design_direction](../.claude/projects/c--Users-user-Desktop-mamoru/memory/feedback_tms_design_direction.md) — 마모루 가이드 100% 추종 X, 트렌드 + 작업효율 우선.
**공통 컴포넌트**: `components/ui/stat-card.tsx` + `components/ui/revenue-dark-card.tsx` (재사용 필수, [feedback_code_dry_no_duplicates](../.claude/projects/c--Users-user-Desktop-mamoru/memory/feedback_code_dry_no_duplicates.md)).

**페이지 그룹별 진행 순서**:
- [x] **1. 대시보드** (2026-05-27 시안 B+ 완료)
- [x] **2. 상담 페이지군** (2026-05-27 완료) — 4탭(전체/매장/출장/톡), 전체 탭 기본값
- [x] **3. 복원수리 페이지군 + /sales 매출 카드** (2026-05-27 완료) — A2 그라데이션 다크 매출 카드
- [x] **4. 주문 페이지군** (2026-05-27 완료) — 상태 7탭 stone-900 + OrderRow 모노크롬
- [x] **5. 판매 상세 페이지군** (2026-05-27 완료, `6a90411`) — sales/[id], sales/new, deliveries, detail panels
- [x] **6. 매입 페이지군** (2026-05-27 완료) — purchasing page/new/[id], purchase-detail-panel
- [x] **7. 고객/상품 페이지군** (2026-05-27 완료) — customers, products(+new/[id]/serials), detail panels
- [x] **8. 시리얼/계약서/설정** (2026-05-27 완료) — contracts, settings(banner/calendar/notification). 시리얼은 이미 준수

**🎉 TMS 톤 통일 전체 완료 (2026-05-27).** 단독 terracotta 0건. 상태색·terracotta-soft 배지는 의도적 유지.
**잔여 미세 조정 (선택)**: 일부 페이지 배경 `bg-stone-50` 미적용분 + 설정 카드 `rounded-lg → rounded-2xl` (낮은 우선순위).

**각 그룹 진행 절차** (재사용):
1. 디자인 모니터(`/design-lab`)에 § 시안 1~2개 추가 → 사장님 비교 (또는 톤 일관성 확인되면 § 단계 생략 가능)
2. 채택안 결정 → 실제 페이지에 적용 + **공통 컴포넌트 재사용 우선** (DRY)
3. 흐름도(`docs/TMS_FLOW_*.md`) 영향 시 동일 커밋에 업데이트
4. `npx tsc --noEmit` 통과 확인
5. 사장님 승인 → push → Vercel 빌드 확인 → 링크 제공
6. design-lab § 즉시 삭제 (운영 룰)

**절대 제약**:
- 회계 RPC (077·078·080·088 등) 미수정
- 데이터 hook 흐름 무수정 — UI 재배치/스타일 변경만
- 매출 합계 결과 = 적용 전과 동일

---

### 3. Phase 1A 상품 상세 카탈로그 검토 (사장님 검토 대기) ⭐

**검토 위치**: `projects/products/catalog/preview.html` 브라우저로 열기

**검토 항목**:
1. §0 매트릭스 — 4가위(블런트/장가위/틴닝/드라이) × 9카드 매핑 정확성
2. §1~7 카드 시각화 — 각 카드 옵션값·SVG·설명
3. 틴닝 옵션값 — 24/26/28/30/32/40 발 / 1·2·3·4 홈 / 15·20·25·30·40% 감모
4. 드라이 옵션 가격 — 스트록 기본 / 정통 +1만 / 멀티(SHIFT ONE) +2만
5. §A 카피 풀 / §B spec — 기존 (변경 없음)

**카드 카탈로그 9개** (`projects/products/catalog/cards/`):
- blade_edge.json (applies_to: blunt/long/dry)
- blade_design.json (S/C/B)
- handle_grip.json (세미/스탠/오프셋)
- handle_camel.json (플랫/카멜)
- grade.json (R/A/E/S)
- thinning_teeth.json / thinning_holes.json / thinning_reduction.json
- dry_cutting_style.json (★ 주문 옵션, 가격 차등)

**사장님 검토 후 액션**:
- "이 옵션 추가/빼기/이름·SVG/가격 수정" 한 마디 → 즉시 반영
- OK → Phase 1B (틴닝 카피 풀) + Phase 1C (드라이 카피 풀)
- 그 다음 → Phase 2 빌더 페이지 개발

---

## 🟡 자동 대기

### 복원수리 직접방문(당일수리) Phase 4 — 솔라피 검수 대기

**현재 상태**: Phase 1~3-B 운영 시작. 알림톡 5종 사양 확정 + 솔라피 검수 사장님 진행 중.

**사장님 외부 작업**:
1. 솔라피 콘솔에서 5종 신규 템플릿 본문 작성 + 변수/버튼 등록
   - `as_visit_booked` / `as_visit_remind` / `as_visit_rescheduled` / `as_visit_cancelled` / `as_visit_completed`
2. 카카오 검수 신청 (1~3 영업일)
3. 검수 통과 후 Make 시나리오 라우터 분기 추가:
   - 접수 알림 시나리오: `as_visit_booked` 분기 1개
   - 상태변경 시나리오: 4종 분기 (`remind/rescheduled/cancelled/completed`)
4. (선택) 리마인드 시점 결정 (A/B/C안 중)

**검수 통과 후 TMS 코드 변경 (클로드에게 요청)**:
- `lib/notification/make-webhook.ts` 5종 추가
- 라우팅: `as_visit_booked` 도 `as_received` 시나리오 합류
- `api/repair/public/submit/route.ts` 직접방문 분기 알림톡
- `api/repair/[id]/route.ts` PATCH 분기 알림톡
- `api/cron/repair-visit-remind/route.ts` 신규 (사장님 리마인드 시점 결정 후)

**상세 사양**: [project_repair_direct_visit](../.claude/projects/c--Users-user-Desktop-mamoru/memory/project_repair_direct_visit.md) + [reference_solapi_templates](../.claude/projects/c--Users-user-Desktop-mamoru/memory/reference_solapi_templates.md) "직접방문 알림톡 5종" 섹션

---

## 📌 OS-20260525-004 복구 (사장님 진행 중)

- ✅ 094/095 마이그레이션 실행 + 백필 정상 (offline_sales 25 = 25)
- ✅ 사장님 수동 발송 ([상담→톡상담]) → 솔라피 정상 도달 (Make `상담만족도&리뷰 20` 분기 OK)
- ⚠️ 자동 cron 단계는 `purchase_review_request` 분기에서 차단 — 위 1번 항목 진단 결과 따라 후속 처리

---

## 범례
- 🟡 = 진행 중 / 대기
- ⚠️ = 사장님 외부 작업 필요
- ⭐ = 사장님 검토 우선순위
