# 복원수리 프로세스 흐름도
> 최종 업데이트: 2026-03-03

---

## 1. 비즈니스 프로세스 흐름

### 전체 파이프라인 (6단계)
```
신규접수 → 입고대기 → 작업중 → 출고대기 → 출고완료 → 배송완료
```

### 방문수거 흐름
```
고객 접수 (page_form.html)
  → GAS doPost: 시트 저장 + as_id 채번(AS-YYYYMMDD-NNN)
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
[고객] ──폼──→ [GAS doPost]
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
| `as_received` (접수완료) | GAS doPost 내부 | GAS → Make webhook |
| `as_cost_notice` (비용안내) | UI "비용안내" 버튼 | UI → POST `/api/repair/[id]/notify` |
| `as_payment_confirmed` (입금확인) | paid_at 플래그 설정 시 | PATCH `/api/repair/[id]` → after() 자동 |
| `as_shipped` (출고완료) | shipped 상태 전환 시 | PATCH `/api/repair/[id]` → after() 자동 |
| `as_satisfaction` (만족도) | 미구현 | — |

### 비용 계산 규칙
```
서비스비 = (마모루 수량 × 10,000) + (타사 수량 × 20,000)
배송비   = 방문수거일 때: 1개=5,000 / 2개=3,000 / 3개+=무료
           직접발송일 때: 0원
합계     = 서비스비 + 배송비
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
| `/api/repair/sync` | POST | GAS → TMS 동기화 |
| `/api/repair/report` | GET | 수리내역 공개 API (CORS, 인증 불필요) |

### 컴포넌트 (15+개)
- 고정 탭 바 6개: 신규접수 / 수거접수필요 / 입고대기 / 진행중 / 출고대기 / 출고완료
- 인라인 액션 칩: 내역서 / 입금확인 / 송장생성 / 포장완료
- PC 마스터-디테일 레이아웃 (lg+)
- 검수 폼 (가위별 7항목: 날끝/중간/안쪽/빗살/텐션/부품/스토퍼)
- 검수 요약 (읽기 전용 테이블)
- 사이드바 액션 카드 (비용+상태버튼+입금/송장)
- 타임라인 (상태 변경 이력)
- 상태 배지 (색상 코딩 + 진행방식 배지)

### 검수 자동 문구 생성
- `lib/repair/inspection-text.ts` — 검수 데이터 분석 → 한국어 작업 설명 자동 생성
- 무뎌짐, 찍힘, 빗살 손상, 텐션 느슨, 부품 교체, 스토퍼 교체 자동 감지

### DB 테이블
- `repairs` — 메인 (as_id, 비용, 상태, 송장, confirmed_at, packed_at, paid_at)
- `repair_inspections` — 가위별 검수 (blade_tip/mid/inner, comb, tension, parts, stopper)
- `repair_history` — 상태 변경 이력

### GAS 스크립트
- `projects/as/Code.gs` — 접수(doPost) + 관리(updateAS_) + 롯데택배(ALPS) + TMS 동기화

### 고객 대면 페이지 (GitHub Pages)
- `page_form.html` — 통합 접수 폼 (마모루+타사)
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
| Make Router에 복원수리 5종 분기 추가 | 상담 17종 Make 연결 완료 후 | 높음 |
| BC 메타데이터 제거 후 솔라피 재검수 대기 | 솔라피 검수 (1~3 영업일) | 높음 |
| 복원수리 E2E 전체 플로우 검증 | Make 분기 + 솔라피 재검수 완료 후 | 높음 |
| 만족도 알림톡 (as_satisfaction) 자동 발송 로직 | 솔라피 검수 | 중간 |
| 사진 업로드 Supabase Storage 연동 | 버킷 생성 | 중간 |
| 주소 수정 시 다음 주소검색 API 연동 | 없음 | 중간 |
| 사진 마킹 (photo-marker) html2canvas | 없음 | 낮음 |
| 수리내역서 자동 생성 (Before/After 웹카드) | 사진 업로드 | 낮음 |

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
