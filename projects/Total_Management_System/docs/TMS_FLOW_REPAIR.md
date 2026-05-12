# 복원수리 프로세스 흐름도
> 최종 업데이트: 2026-05-12 — **복원수리 매출 3채널 통합 (A 접수 + B 판매RS + C 납품RS) + 거래처별 단가 + 판매입력 복원수리 모드** (코드 완료, push 대기)
>
> **마스터 문서**: [TMS_SYSTEM_ARCHITECTURE.md](TMS_SYSTEM_ARCHITECTURE.md) §3 참조

## 2026-05-12 — 복원수리 매출 3채널 통합 + 입력 경로 정리

복원수리 매출이 3개 채널에서 발생 → 대시보드·회계 모두 통합 집계:
- **A채널 (접수시스템 `repairs`)**: 고객이 접수페이지로 가위 보냄 → 입고·검수·수리·출고. 단가는 마모루 1만 / 타사 2만 고정(대시보드) 또는 service_cost+shipping_fee 실제 금액(회계, paid_at 기준).
- **B채널 (판매시스템 `offline_sale_items` category='RS')**: 매장에서 즉석 복원수리 판매. 입력: `/sales/new` "복원수리" 모드 (마모루 N자루 / 타사 M자루 + 단가, 기본 1만/2만). product_name 고정 "복원수리 (마모루)" / "복원수리 (타사)" → `includes('타사')` 로 마모루/타사 분류. parent 판매의 customer_type 이 dealer/academy 면 B2B 버킷으로.
- **C채널 (납품 `delivery_items` category='RS')**: B2B 거래처 복원수리. 입력: `/deliveries` "+B2B수리" 모드. 마모루/타사 구분 없음(자루당 거래처 단가). 거래처 마스터의 `default_repair_price` 가 있으면 거래처 선택 시 단가 자동 채움(없으면 기본 8천).
- 대시보드 `monthRepairAmount` = A + B + C 전체, `monthRepairCount` = 세 채널 자루(quantity) 합. `monthRepairMamoru/Other` = A + B채널 B2C분, `monthRepairB2B` = B채널 B2B분 + C채널 전체.
- 회계 리포트(2-D): `repair_sales.total` = A(repairs, paid_at, 실제금액) + B(offline RS) + C(delivery RS), 채널별 분해 표시.
- 거래처 단가: `customers.default_repair_price` (마이그레이션 079). 고객 상세 화면(딜러/아카데미)에서 입력.

## 2026-04-30 (심야) — 고객 자동 매칭/생성

복원수리 접수 시 phone 기준 자동 매칭/생성:
- `/api/repair/public/submit` (고객 폼) → `lib/customer/match-or-create.ts` 호출 → repairs.customer_id 자동 연결
- `/api/repair` POST (관리자 수기) → body.customer_id 없을 때만 자동 매칭 (사장님 customer-autocomplete 선택은 그대로)

같은 phone으로 상담/복원수리/판매 재접수 시 동일 customer_id 매칭 → 거래 이력 통합 + 단골 자동 인지.

상세: `docs/TMS_FLOW_CONSULTATION.md` "2026-04-30 심야" 섹션

---

## 리뷰 요청 분기 (2026-04-29 추가)

배송완료(`delivered`) 진입 시 후기 알림톡 발송 정책. `system_settings.review.auto_request_on_completion` 토글로 두 모드 양립.

- **Mode A (default OFF)**: 자동 발송 X, 사장님 수동만
- **Mode B (ON)**: 약속 X 고객만 자동, 약속 ✓ 고객은 항상 수동

3 timestamp: `review_promised_at` / `review_request_sent_at` / `review_submitted_at`. reviews/submit가 `repairs.as_id` 매칭으로 review_submitted_at 자동 기록. 상세 패널의 "리뷰 관리" 카드와 `/reviews` "약속 대기" 탭에서 추적.

---

---

## 1. 비즈니스 프로세스 흐름

### 전체 파이프라인 (6단계)
```
신규접수 → 입고대기 → 작업중 → 출고대기 → 출고완료 → 배송완료
```

### 방문수거 흐름
```
고객 접수 (page_form.html)
  → Vercel API: /api/repair/public/submit → Supabase 저장 + as_id 채번(AS-YYYYMMDD-NNN)
  → Make 알림톡 (as_received) + TMS 동기화
  → [status: intake] (신규접수)
  → (관리자) 접수확인 → confirmed_at 기록
  → (관리자) 수거접수 완료 → [status: pickup_scheduled] (입고대기)
  → 입고 & 검수 + 비용안내 → [status: cost_notified] (작업중)
    → 알림톡 (as_cost_notice) — UI 버튼 → sendNotify API
  → (고객) 입금 → (관리자) 입금확인 → paid_at 기록
    → 알림톡 (as_payment_confirmed) — PATCH API 내부 자동
  → 수리 진행 → [status: repairing] (작업중)
  → 포장 완료 → packed_at 기록
  → 송장 생성 (ALPS) → [status: ready_to_ship] (출고대기)
  → 출고 완료 → [status: shipped] (출고완료)
    → 알림톡 (as_shipped) — PATCH API → getAutoNotifyTemplate 자동
  → 배송 완료 → [status: delivered] → [status: completed]
  → 만족도 알림톡 (as_satisfaction) — 미구현
  ※ 모든 단계에서 취소 가능 → [status: cancelled]
    → 알림톡 (as_cancelled) — PATCH API → getAutoNotifyTemplate 자동
```

### 직접발송 흐름
```
고객 접수 → [intake]
  → (관리자) 접수확인 → confirmed_at
  → 입고 & 검수 + 비용안내 → [cost_notified] (pickup_scheduled 생략)
  → 이하 방문수거와 동일
```

### 상태 전이 규칙
```
intake → [pickup_scheduled, cost_notified, cancelled]
pickup_scheduled → [cost_notified, cancelled]
cost_notified → [repairing, cancelled]
repairing → [ready_to_ship, cancelled]
ready_to_ship → [shipped]
shipped → [delivered]
delivered → [completed]
completed → (terminal)
cancelled → (terminal)

* paid_at: 파이프라인과 독립된 플래그 (어느 상태에서든 입금확인 가능)
* proceed_type 필터: 방문수거는 pickup_scheduled 필수, 직접발송은 생략
```

---

## 2. 시스템 연동 흐름

```
[고객] ──폼──→ [Vercel API: /api/repair/public/submit]
                  │
         ┌───────┼───────────┐
         ▼       ▼           ▼
   [Google 시트] [Make]   [TMS /api/repair/sync]
                  │           │
                  ▼           ▼
             [Solapi]    [Supabase DB]
             (알림톡)         │
                              ▼
                        [TMS React UI]
                              │
            ┌─────────────────┼────────────────┐
            ▼                 ▼                ▼
   [검수/비용안내]     [송장생성(ALPS)]    [알림톡 발송]
   PATCH /repair/[id]  POST /repair/[id]/ship  POST /repair/[id]/notify
                              │                     또는 PATCH 자동
                              ▼
                        [롯데택배 ALPS API]
                        (운송장 12자리 생성)
```

### 알림톡 발송 구조 (4가지 트리거)

| 템플릿 | 트리거 | 경로 |
|--------|--------|------|
| `as_received` (접수완료) | Vercel submit route | Vercel → Make webhook |
| `as_cost_notice` (비용안내) | UI "비용안내" 버튼 | UI → POST `/api/repair/[id]/notify` |
| `as_payment_confirmed` (입금확인) | paid_at 플래그 설정 시 | PATCH `/api/repair/[id]` → after() 자동 |
| `as_shipped` (출고완료) | shipped 상태 전환 시 | PATCH `/api/repair/[id]` → after() 자동 |
| `as_cancelled` (취소안내) | cancelled 상태 전환 시 | PATCH `/api/repair/[id]` → after() 자동 |
| `as_satisfaction` (만족도) | 미구현 | — |

### Make 웹훅 URL 분기 (2026-03-05)
| 웹훅 | 환경변수 | 대상 템플릿 |
|------|----------|-------------|
| 상담 알림톡 | `MAKE_WEBHOOK_URL` | confirmed, cancelled, suggest 등 상담 17종 |
| 복원수리 접수/취소 | Vercel API | as_received (Vercel → Make webhook) |
| 복원수리 상태변경 | `MAKE_REPAIR_WEBHOOK_URL` | as_cost_notice, as_payment_confirmed, as_shipped, as_cancelled, as_satisfaction |

### 웹훅 payload 주요 필드
| 필드 | 값 | 용도 |
|------|-----|------|
| `id` | repair.as_id (AS-YYYYMMDD-NNN) | Make `#{as_uid}` → 수리내역서 링크 |
| `tracking` | 송장번호 | 배송조회 |
| `courier` | 롯데택배 (고정) | 배송조회 버튼 활성화 |
| `as_amount` / `shipping_amount` / `total_amount` | 금액 | 비용 안내 |
| `pickup_date` | `YYYY년 MM월 DD일 (X요일)` 포맷 | 방문수거 접수 알림톡 수거예정일 (as_received) · 직접발송은 빈 문자열 |
| `delivery_method` | 문앞보관 / 매장카운터 / 직접전달 | 방문수거 수령 방식 (as_received) |

### 비용 계산 규칙
```
서비스비 = (마모루 수량 × 10,000) + (타사 수량 × 20,000)
배송비   = 방문수거일 때: 1개=6,000 / 2개=3,000 / 3개+=무료
           직접발송일 때: 1개=3,000 / 2개+=무료
합계     = 서비스비 + 배송비
※ 실제 우체국 수거+발송 비용 8,000원 대비 할인 적용
```

---

## 3. 구현 완료 ✅

### API Routes (7개)
| 엔드포인트 | 메서드 | 기능 |
|------------|--------|------|
| `/api/repair` | GET | 목록 조회 (상태그룹 필터/검색/페이징) |
| `/api/repair` | POST | 신규 접수 생성 |
| `/api/repair/[id]` | GET | 단건 + 검수 + 이력 조회 |
| `/api/repair/[id]` | PATCH | 상태/필드 변경 + 이력 + 자동 알림톡 |
| `/api/repair/[id]/inspect` | POST/PUT | 검수 데이터 저장/수정 |
| `/api/repair/[id]/notify` | POST | 수동 알림톡 발송 |
| `/api/repair/[id]/ship` | POST/DELETE | 송장 생성(ALPS) / 취소 |
| `/api/repair/sync` | POST | 외부 동기화 (레거시, 사용 빈도 낮음) |
| `/api/repair/report` | GET | 수리내역 공개 API (CORS, 인증 불필요) |

### 컴포넌트 (15+개)
- 고정 탭 바 6개: 신규접수 / 수거접수필요 / 입고대기 / 진행중 / 출고대기 / 출고완료
- **인라인 퀵 액션** (03-26 리모델): 탭별 특화 — 접수확인/수거접수/입금확인/송장생성/출고완료
- **상단 요약 카드 4개** (03-26→03-30): 신규접수(blue) / 진행중(amber) / 미입금(red) / 3일경과(orange) — **클릭 시 탭 이동 + 교차 필터 활성** (토글)
- **useRepairTabData 기반 리스트** (03-26): 7탭→6탭 전환, 비즈니스 로직 쿼리 활용
- **isLg conditional rendering** (03-26): CSS hidden→조건부 렌더링 (SlidePanel 이슈 해결)
- PC 마스터-디테일 레이아웃 (lg+) — 좌측 480px 고정 + 우측 flex-1 상세 모니터
- 검수 폼 (가위별 7항목: 날끝/중간/안쪽/빗살/텐션/부품/스토퍼)
- 검수 요약 (읽기 전용 테이블)
- 사이드바 액션 카드 (비용+상태버튼+입금/송장)
- 타임라인 (상태 변경 이력)
- 상태 배지 (색상 코딩 + 진행방식 배지)
- **탭별 특화 정보 칩** (03-26): 수거필요→주소, 진행중→입금상태, 출고대기→송장/포장, 출고완료→배송상태
- **완료/취소 시각 구분** (03-26): green border-left / opacity+취소선 / 미입금=orange border / 미입금+3일경과=red border
- **입금확인 알림톡 선택** (03-30): 체크박스로 발송/미발송 분리 (skip_notify 플래그)
- **확인 모달 전체 적용** (03-29): 비용안내/입금확인/출고완료/송장취소/수리취소 — ConfirmModal 공통 컴포넌트

### 검수 자동 문구 생성
- `lib/repair/inspection-text.ts` — 검수 데이터 분석 → 한국어 작업 설명 자동 생성
- 무뎌짐, 찍힘, 빗살 손상, 텐션 느슨, 부품 교체, 스토퍼 교체 자동 감지

### DB 테이블
- `repairs` — 메인 (as_id, 비용, 상태, 송장, confirmed_at, packed_at, paid_at)
- `repair_inspections` — 가위별 검수 (blade_tip/mid/inner, comb, tension, parts, stopper)
- `repair_history` — 상태 변경 이력

### ~~GAS 스크립트~~ (제거 완료 — 2026-04-01)
- ~~`projects/as/Code.gs`~~ — Vercel API로 완전 대체, 보관처리 예정

### 고객 대면 페이지 (GitHub Pages)
- `page_form.html` — 통합 접수 폼 (마모루+타사)
- `page_guide.html` — 복원수리 안내 (아임웹 iframe용, 6탭: 마모루복원수리/과정안내/소요시간/비용안내/포장방법/QnA, 라이트모드+다크 intro, CTA 통합접수)
- `page_as_guide.html` — 복원수리 안내 (알림톡 링크용, 모바일 전용 4탭)
- `page_as_report.html` — 수리내역 조회 (TMS API 연동)

---

## 4. 완료된 외부 연동 ✅ (2026-03-03)

| 항목 | 완료일 |
|------|--------|
| 솔라피 복원수리 5종 템플릿 등록 + 검수 승인 | 2026-03-03 |
| GAS Script Properties 설정 (TMS_REPAIR_SYNC_URL, CRON_SECRET) | 2026-02-28 |
| 복원수리 접수 → GAS → TMS 동기화 테스트 통과 | 2026-02-28 |

## 5. 미완료 ❌

| 항목 | 의존성 | 우선순위 |
|------|--------|----------|
| ~~Make Router에 복원수리 6종 분기 추가~~ | ~~완료~~ | ~~✅ 03-12~~ |
| ~~솔라피 as_cancelled 취소안내 템플릿~~ | ~~완료~~ | ~~✅ 03-03~~ |
| ~~BC 메타데이터 제거 후 솔라피 재검수~~ | ~~완료~~ | ~~✅ 03-19~~ |
| ~~사진 업로드 Supabase Storage 연동~~ | ~~완료~~ | ~~✅ 03-24~~ |
| 복원수리 E2E 전체 플로우 검증 | Make 분기 + 솔라피 재검수 완료 후 | 높음 |
| 만족도 알림톡 (as_satisfaction) 자동 발송 로직 | 솔라피 검수 | 중간 |
| 주소 수정 시 다음 주소검색 API 연동 | 없음 | 중간 |
| 사진 마킹 (photo-marker) html2canvas | 없음 | 낮음 |
| 수리내역서 자동 생성 (Before/After 웹카드) | 사진 업로드 ✅ | 낮음 |

---

## 6. 핵심 파일 맵

### TMS API
| 파일 | 설명 |
|------|------|
| `app/src/app/api/repair/route.ts` | GET/POST 목록/생성 |
| `app/src/app/api/repair/[id]/route.ts` | GET/PATCH 단건 + 상태머신 |
| `app/src/app/api/repair/[id]/inspect/route.ts` | POST/PUT 검수 데이터 |
| `app/src/app/api/repair/[id]/notify/route.ts` | POST 알림톡 발송 |
| `app/src/app/api/repair/[id]/ship/route.ts` | POST/DELETE 송장 |
| `app/src/app/api/repair/sync/route.ts` | POST GAS 동기화 |
| `app/src/app/api/repair/report/route.ts` | GET 공개 수리내역 API |

### TMS UI
| 파일 | 설명 |
|------|------|
| `app/src/app/(dashboard)/repairs/page.tsx` | 복원수리 메인 (탭 바+목록) |
| `app/src/app/(dashboard)/repairs/dashboard/page.tsx` | 복원수리 대시보드 |
| `app/src/app/(dashboard)/repairs/[id]/page.tsx` | 복원수리 상세 (모바일) |
| `app/src/components/repairs/repair-tab-bar.tsx` | 고정 탭 바 6개 |
| `app/src/components/repairs/repair-list.tsx` | 마스터 목록 |
| `app/src/components/repairs/repair-detail-panel.tsx` | PC 상세 패널 |
| `app/src/components/repairs/repair-detail-card.tsx` | 고객/접수 정보 카드 |
| `app/src/components/repairs/repair-action-chips.tsx` | 인라인 액션 칩 |
| `app/src/components/repairs/inspection-form.tsx` | 검수 입력 폼 |
| `app/src/components/repairs/inspection-summary.tsx` | 검수 요약 |
| `app/src/components/repairs/repair-timeline.tsx` | 이력 타임라인 |
| `app/src/components/repairs/sidebar-action-card.tsx` | 사이드바 액션 |
| `app/src/components/repairs/repair-status-badge.tsx` | 상태 배지 |
| `app/src/components/repairs/tabs/*.tsx` | 탭별 구현 (6개) |

### TMS Lib
| 파일 | 설명 |
|------|------|
| `app/src/lib/repair/transitions.ts` | 상태 전이 규칙 (v3) |
| `app/src/lib/repair/cost-calculator.ts` | 비용 계산 |
| `app/src/lib/repair/inspection-text.ts` | 검수→작업문구 자동생성 |
| `app/src/lib/repair/sync.ts` | GAS→DB 동기화 로직 |
| `app/src/lib/notification/make-webhook.ts` | Make webhook (상담과 공유) |
| `app/src/hooks/use-repairs.ts` | React Query 훅 8개 |
| `app/src/hooks/use-repair-tabs.ts` | 탭별 쿼리+카운트 훅 |

### GAS (Google Apps Script)
| 파일 | 설명 |
|------|------|
| `projects/as/Code.gs` | 접수/관리/ALPS/TMS동기화 |

### 롯데택배 ALPS
| 파일 | 설명 |
|------|------|
| `app/src/lib/lotte/client.ts` | ALPS API (book/cancel/track) |
| `app/src/lib/lotte/types.ts` | 타입 정의 |

### DB 마이그레이션
| 파일 | 설명 |
|------|------|
| `app/supabase/migrations/005_repair_paid_at.sql` | paid_at 플래그 |
| `app/supabase/migrations/006_repair_confirmed_packed.sql` | confirmed_at, packed_at |

### 고객 대면 페이지
| 파일 | 설명 |
|------|------|
| `projects/as/page_form.html` | 통합 접수 폼 |
| `projects/as/page_as_report.html` | 수리내역 조회 |
