# MAMORU 시스템 구축 — TODO

> 최종 수정: 2026-05-27 — **TMS 톤 통일 4그룹 완료** (대시보드/상담/복원수리+sales/주문) + **공통 컴포넌트 추출** (StatCard + RevenueDarkCard) + **리뷰 약속 유형/subtype 시스템** (094/095 마이그레이션) + **IA 모순 해결** (약속 ON 시 모달 우회) + **알림톡 진단 사전 안내 룰 박제**.
>
> 완료 이력은 모두 git history(`git log --oneline`) 참조. 본 파일은 진행 중·대기만 유지.

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
- [ ] **5. 판매 상세 페이지군** (`/sales/[id]`, `/sales/new`, `/deliveries`, `/deliveries/[id]`) ← 다음
- [ ] **6. 매입 페이지군** (`/purchasing` 발주/입고/상세 — purchase-detail-panel.tsx)
- [ ] **7. 고객/상품 페이지군** (`/customers`, `/products`, `/categories`)
- [ ] **8. 시리얼/계약서/설정** (`/serials`, `/contracts`, `/settings`)

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
