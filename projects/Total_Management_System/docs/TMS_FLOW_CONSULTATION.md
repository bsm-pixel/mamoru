# 상담관리 프로세스 흐름도
> 최종 업데이트: 2026-03-03 (알림톡 검수 승인 + BC 메타데이터 제거 + Make Router 진행)

---

## 1. 비즈니스 프로세스 흐름

### 매장방문 (store_visit)
```
고객 접수 (아임웹 폼)
  → GAS: 시트 저장 + 슬롯 차감 + 구글 캘린더 등록 + 접수완료 알림톡
  → TMS Supabase 자동 동기화
  → [status: confirmed] (즉시 확정)
  → (고객) 셀프서비스 변경/취소 가능
  → [리마인드] D-1(24H) + D-0(2H) 자동 알림톡
  → 방문 완료 → [status: completed]
```

### 출장요청 (field_request)
```
고객 접수 (아임웹 폼)
  → GAS: 시트 저장 + 접수 알림톡(request)
  → TMS Supabase 자동 동기화
  → [status: pending_admin]
  → (관리자) 딜러 배정 → [status: assigned]
  → (관리자) 시간 제안 (최대 3슬롯) → [status: suggested] + 제안 알림톡
  → (고객) 시간 선택 → [status: confirmed] + 확정 알림톡
       또는 재요청 → [status: reschedule_requested]
       또는 변경요청 → [status: change_requested]
  → [리마인드] D-1(24H) + D-0(2H) 자동 알림톡
  → (관리자) 출장 지연 시 → 지연안내 알림톡 (delay)
  → 출장 완료 → [status: completed]
```

### 톡상담 (talk_consult)
```
고객 접수 (카카오톡 유입 → 해피톡)
  → GAS: 시트 저장 + 접수 알림톡(talk_received)
  → TMS Supabase 자동 동기화
  → [status: pending_admin]
  → (관리자) "상담 시작" → [status: in_progress] + talk_ready 알림톡
  → 상담 종료 → [status: completed]
```

### 상태 전이 규칙
```
매장방문: confirmed → [completed, on_hold, cancelled]
         on_hold → [confirmed, cancelled]

출장요청: pending_admin → [suggested, confirmed, on_hold, cancelled]
         suggested → [confirmed, reschedule_requested, on_hold, cancelled]
         reschedule_requested → [suggested, confirmed, on_hold, cancelled]
         change_requested → [suggested, confirmed, on_hold, cancelled]
         confirmed → [completed, on_hold, cancelled]
         on_hold → [confirmed, cancelled]

톡상담:   pending_admin → [in_progress, on_hold, cancelled]
         in_progress → [completed, on_hold, cancelled]
         on_hold → [in_progress, cancelled]
```

---

## 2. 시스템 연동 흐름

```
[고객]
  │
  ▼
[아임웹 폼] ──POST──→ [GAS Code.gs]
                          │
                ┌─────────┼─────────────────┐
                ▼         ▼                 ▼
         [Google 시트]  [Google 캘린더]   [Make Webhook]
                │                           │
                ▼                           ▼
         [supabase-sync.gs]            [Solapi 알림톡]
                │
                ▼
         [TMS /api/consultation/sync]
                │
                ▼
         [Supabase DB]
                │
                ▼
         [TMS React UI] ──액션──→ [TMS API] ──→ [GAS (캘린더/슬롯)]
                                     │
                                     └──→ [Make Webhook → Solapi]
```

### 연동별 역할
| 시스템 | 역할 | 방향 |
|--------|------|------|
| GAS Code.gs | 접수 처리, 캘린더, 슬롯, Make 호출, TMS 동기화 | 아임웹 → GAS → TMS |
| Make Webhook | 알림톡 발송 허브 (17종 분기) | TMS/GAS → Make → Solapi |
| Solapi | 카카오 알림톡 실제 발송 | Make → Solapi → 고객 |
| Google Calendar | 매장/출장 일정 관리 | GAS ↔ Calendar |
| Supabase | 상담 데이터 영구 저장 | GAS → TMS → Supabase |

---

## 3. 구현 완료 ✅

### API Routes (7개)
| 엔드포인트 | 메서드 | 기능 |
|------------|--------|------|
| `/api/consultation` | GET | 목록 조회 (필터/페이징/검색) |
| `/api/consultation` | POST | 신규 상담 생성 |
| `/api/consultation/[id]` | GET | 단건 + 이력 조회 |
| `/api/consultation/[id]` | PATCH | 상태/정보 변경 + 이력 기록 + 자동 알림톡 |
| `/api/consultation/assign` | POST | 딜러 배정 |
| `/api/consultation/suggest` | POST | 출장 시간 제안 (GAS 연동) |
| `/api/consultation/delay` | POST | 출장 지연 안내 (GAS 연동) |
| `/api/consultation/notify` | POST | 수동 알림톡 발송 |
| `/api/consultation/sync` | POST | GAS → TMS 데이터 동기화 |

### 컴포넌트 (10+개)
- 매장방문 리스트 (3탭: 확정/취소/완료)
- 출장요청 리스트 (4탭: 신규/제안/재요청/확정) + 카카오맵 양방향 연동
- 톡상담 리스트 (4탭: 신규/진행중/완료/취소)
- 일정 달력 (초록: 매장, 보라: 출장)
- 시간 제안 모달 (최대 3슬롯)
- 일정 변경 모달
- 보류 사유 모달
- 모바일 출장 일별 뷰
- 상담 대시보드 (통계 + 오늘일정)

### 알림톡 (17종 — Make webhook 경유)
```
매장방문 (5종): confirmed, cancelled, rescheduled, remind24, remind2
출장요청 (9종): request, suggest, field_confirmed, field_cancelled,
                field_rescheduled, field_remind_24h, field_remind_2h,
                field_delayed, change_request_received
톡상담   (2종): talk_received, talk_ready
공통     (1종): as_received (복원수리 접수)
```

### 알림톡 버튼 설계 (2026-03-03 확정)
```
WL 버튼 (웹링크): 6개 — 일정확인/변경, 일정 선택하기, 복원수리 안내 확인
  → URL에 https:// 프로토콜 필수 (솔라피 규칙)
  → change_request_link: GAS/TMS 모두 프로토콜 미포함 → 솔라피 템플릿에서 https://#{변수} 형태

BC 버튼 (상담톡전환): 12개 — 1:1 문의하기
  → chatExtra 메타데이터: 제거 (한글 문자 불가 이슈 — 3080 에러)
  → 고객 식별: 해피톡 진입 시 사전 입력 폼(성함/연락처)으로 대체
```

### DB 테이블
- `consultations` — 상담 메인 (003 + 004 마이그레이션)
- `consultation_history` — 상태 변경 이력
- `dealers` — 딜러 목록 (regions, calendar_id)

### GAS 스크립트
- `projects/consulting/Code.gs` — 접수, 캘린더, 취소, 재일정 등
- `projects/consulting/supabase-sync.gs` — TMS 동기화

### 고객 대면 페이지 (GitHub Pages)
- `page_suggest.html` — 출장 일정 선택 (캘린더 UI)
- `page_change_request.html` — 예약 변경/취소 셀프서비스

---

## 4. 완료된 외부 연동 ✅ (2026-03-03)

| 항목 | 완료일 |
|------|--------|
| 솔라피 17종 전체 검수 제출 + 승인 | 2026-03-03 |
| BC 버튼 chatExtra 한글 불가 발견 → 메타데이터 전체 제거 결정 | 2026-03-03 |
| TMS 일정변경(reschedule) 알림톡 change_request_link 누락 수정 | 2026-03-03 |
| TMS 일정변경 시 template 자동 분기 (rescheduled / field_rescheduled) | 2026-03-03 |

## 5. 미완료 ❌

| 항목 | 의존성 | 우선순위 |
|------|--------|----------|
| Make Router 17개 분기 연결 | 진행 중 (수동 작업) | 높음 |
| BC 메타데이터 제거 후 솔라피 재검수 대기 (1~3 영업일) | 솔라피 검수 | 높음 |
| 재검수 승인 후 E2E 알림톡 테스트 | 솔라피 재검수 + Make 완료 후 | 중간 |
| Make → 솔라피 직접 호출 전환 (비용 절감) | 검수 안정화 후 | 낮음 |
| 솔라피: confirmed/rescheduled/field_confirmed 템플릿에 "일정확인/변경" WL 버튼 추가 → 재검수 | 솔라피 관리자 작업 | 중간 |

> 참고: 코드/DB/UI는 100% 완성. 미완료는 모두 외부 서비스(솔라피/Make) 설정 작업.

---

## 6. 핵심 파일 맵

### TMS API
| 파일 | 설명 |
|------|------|
| `app/src/app/api/consultation/route.ts` | GET/POST 상담 |
| `app/src/app/api/consultation/[id]/route.ts` | GET/PATCH 단건 |
| `app/src/app/api/consultation/assign/route.ts` | 딜러 배정 |
| `app/src/app/api/consultation/suggest/route.ts` | 시간 제안 |
| `app/src/app/api/consultation/delay/route.ts` | 지연 안내 |
| `app/src/app/api/consultation/notify/route.ts` | 알림톡 발송 |
| `app/src/app/api/consultation/sync/route.ts` | GAS 동기화 |

### TMS UI
| 파일 | 설명 |
|------|------|
| `app/src/app/(dashboard)/consultations/page.tsx` | 상담관리 메인 (3탭) |
| `app/src/app/(dashboard)/consultations/dashboard/page.tsx` | 상담 대시보드 |
| `app/src/app/(dashboard)/consultations/[id]/page.tsx` | 상담 상세 |
| `app/src/components/consultations/store-visit-list.tsx` | 매장방문 리스트 |
| `app/src/components/consultations/field-request-list.tsx` | 출장요청 리스트 |
| `app/src/components/consultations/field-request-map.tsx` | 카카오맵 |
| `app/src/components/consultations/talk-consult-list.tsx` | 톡상담 리스트 |
| `app/src/components/consultations/schedule-calendar.tsx` | 일정 달력 |
| `app/src/components/consultations/suggest-time-modal.tsx` | 시간 제안 모달 |
| `app/src/components/consultations/reschedule-modal.tsx` | 일정 변경 모달 |
| `app/src/components/consultations/hold-reason-modal.tsx` | 보류 사유 모달 |

### TMS Lib
| 파일 | 설명 |
|------|------|
| `app/src/lib/consultation/sync.ts` | GAS→DB 동기화 로직 |
| `app/src/lib/consultation/transitions.ts` | 상태 전이 규칙 |
| `app/src/lib/notification/make-webhook.ts` | Make webhook 발송 |
| `app/src/hooks/use-consultations.ts` | React Query 훅 10개 |
| `app/src/hooks/use-dashboard-stats.ts` | 대시보드 통계 훅 |

### GAS (Google Apps Script)
| 파일 | 설명 |
|------|------|
| `projects/consulting/Code.gs` | 상담 메인 GAS (캘린더+슬롯+Make) |
| `projects/consulting/supabase-sync.gs` | TMS 동기화 |
| `projects/consulting/products.js` | 제품 데이터 (진단 폼용) |

### DB 마이그레이션
| 파일 | 설명 |
|------|------|
| `app/supabase/migrations/003_consultations.sql` | consultations + history + dealers |
| `app/supabase/migrations/004_consultation_renewal.sql` | talk_consult + on_hold/in_progress/completed + 좌표 |

### 고객 대면 페이지
| 파일 | 설명 |
|------|------|
| `projects/consulting/page_suggest.html` | 출장 일정 선택 캘린더 |
| `projects/consulting/page_change_request.html` | 예약 변경/취소 셀프서비스 |
| `projects/consulting/page_reschedule.html` | 레거시 리다이렉트 (→ page_suggest) |
