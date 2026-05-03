# 상담관리 프로세스 흐름도
> 최종 업데이트: 2026-05-03 — **수동 일정 확정 버튼 (079)** + 달력 관리(078) + 고객 자동 매칭(072)

---

## 2026-05-03 (079) — 수동 일정 확정 버튼 신설

### 변경 배경
사장님 1인 운영 — DM/유선으로 협의된 일정을 즉시 시스템에 박제할 수단 필요. 기존 "시간 제안 → 고객 선택" 흐름 외에 **사장님 직접 확정 진입점** 추가.

### 사용 흐름
```
이미 접수된 건 (pending_admin / suggested / reschedule_requested 등)
  → 사장님 DM/유선 협의
  → TMS 우측 상세 패널 → "수동 일정 확정" 버튼
  → 모달 (날짜 + 시간 입력)
  → 확정 처리

자동 후속 작업 (기존 PATCH 로직 재사용):
  → 상태 confirmed로 전환
  → consultation_history 이력 기록
  → Google Calendar 갱신 (이전 상태 이벤트 정리 + 새 이벤트)
  → 알림톡 자동 발송 (field_confirmed/confirmed 템플릿)
```

### 신규 파일
- `app/src/components/consultations/manual-confirm-modal.tsx` — 모달 + PATCH 호출

### 수정
- `app/src/components/consultations/consultation-detail-panel.tsx` — 버튼 + 모달 통합

### 미수정 (회귀 안전)
- `/api/consultation/[id]/route.ts` — 기존 PATCH 그대로 활용 (status + visit_date + visit_time 모두 처리)
- `lib/google/calendar-sync.ts` 0 변경
- `lib/consultation/transitions.ts` 0 변경 (모든 active 상태 → confirmed 이미 허용)
- `CreateConsultationModal` (일정 수동 등록), `SuggestTimeModal` (시간 제안) 별개로 존재

### 사장님 룰 준수
- closed_dates 검증 추가 X (사장님 측 흐름 = 항상 유동, `feedback_consultation_blackout_rule.md`)
- 노출 조건: `consultation_type !== 'talk_consult'` AND `status NOT IN ('completed', 'cancelled', 'in_progress')`

### 기존 "일정 수동 등록"과의 분담
| | 일정 수동 등록 | 수동 일정 확정 (079) |
|--|---|---|
| 진입점 | 상담관리 상단 버튼 | 우측 상세 패널 |
| 대상 | 접수 X 새 고객 | 기존 접수 건 |
| DB | INSERT | UPDATE |
| API | `admin-create` | `[id]` PATCH |

---

## 2026-05-02 (078) — 달력 관리 + 휴무 SSOT 통합

### 변경 배경
사장님 결정: 휴무 관리를 한 화면에서 통합. 정기 휴무 요일(매주) + 임시 휴무일(특정 날짜)을 분리된 두 곳(설정 페이지 + 별도 SQL)에서 관리하던 것을 **달력 관리** 단일 화면으로 일원화.

### 핵심 룰 (불변)
**막힘은 고객 셀프 예약 흐름에만 적용. 사장님 측 흐름은 항상 유동.**

| 흐름 | 입력자 | closed_dates / disabled_weekdays 검증 |
|------|--------|--------------------------------------|
| 고객 셀프 예약 폼 (`page_form.html`) | 고객 | ✅ 적용 (`/api/consultation/public/slots`, `/settings`) |
| 일정수동등록 (`CreateConsultationModal`) | 사장님 | ❌ 무시 — 단순 input, 항상 유동 |
| 시간 제안 (`SuggestTimeModal`) | 사장님 | ❌ 무시 — 단순 input, 항상 유동 |
| 알림톡 / 리마인더 / Google Calendar | 시스템 | ❌ 무시 — status 변화 trigger만 |

### 신규 파일
- `app/src/app/(dashboard)/consultations/calendar/page.tsx` — 4개월 달력 + 휴무 토글 모달
- `app/src/app/api/consultation/blackouts/route.ts` — closed_dates GET/POST/DELETE
- `app/src/app/api/consultation/settings/route.ts` — consultation_settings GET/PATCH (disabled_weekdays 등)
- `app/src/hooks/use-blackouts.ts` — useBlackouts / useCreateBlackout / useDeleteBlackout / useConsultationSettings / useUpdateDisabledWeekdays

### 수정 (SSOT 통합)
- `components/settings/consultation-settings.tsx`: 휴무 요일 + 특별 휴무일 UI **삭제** → "달력 관리로 이동" 링크 카드로 대체
- `lib/utils/constants.ts`: NAV_GROUPS 상담 그룹에 "달력 관리" 메뉴 추가
- `components/layout/sidebar.tsx` + `mobile-nav.tsx`: CalendarOff 아이콘 추가, matchPrefix 분기

### DB
- `closed_dates` 테이블 (이미 존재, 041_consultation_settings.sql) 그대로 사용 — 마이그레이션 불필요
- `consultation_settings.disabled_weekdays` (이미 존재) 그대로 사용

### 흐름
```
사장님 → /consultations/calendar
  ├─ 정기 휴무 요일 토글 (월~일 7개)
  │   → PATCH /api/consultation/settings { disabled_weekdays: [...] }
  │   → consultation_settings UPSERT
  │
  └─ 날짜 클릭 → 모달
      ├─ 그 날 예약 list 표시 (매장/출장)
      ├─ 정기 휴무 요일 안내 (해당하면)
      ├─ 사유 입력 (선택, 고객 비공개)
      └─ "이 날짜 막기" → POST /api/consultation/blackouts → closed_dates UPSERT
         "휴무 해제" → DELETE /api/consultation/blackouts → closed_dates DELETE

자동 반영:
  → /api/consultation/public/settings 응답에 자동 포함
  → 고객 폼이 30초 TTL 캐시로 1분 안에 반영
  → 사장님 측 모달은 검증 0 (영향 없음)
```

### 휴무 사유 노출 정책
- DB: `closed_dates.reason` (TEXT, 사장님 자유 입력)
- 사장님 UI: 모달에서 표시
- 고객 폼: API 응답에서 `reason` 필드 자동 제외 (`/api/consultation/public/settings/route.ts:46` — `(d) => d.date`만 추출)
- → 고객은 사유 못 봄. 사장님이 "결혼식" 같은 사적 메모도 자유 입력 가능

### 회귀 안전
- consultations 테이블 schema 0 변경
- 기존 예약(confirmed/suggested 등): 정기 휴무 토글에 영향 받지 X (slot 생성 단계만 검증)
- 알림톡/리마인더/Google Calendar: closed_dates 무관 — 정상 동작 유지

---

## 2026-04-30 (심야) — 고객 자동 매칭/생성 (phone 기반 SSOT)

### 변경 배경
사장님 결정: 상담/복원수리 접수 시 phone 기준으로 고객 자동 매칭/생성. 같은 고객이 재접수 시 단골 자동 인지 + 거래 이력 통합. 사장님 입력 부담 절감.

### Fix
**신규 helper**: `app/src/lib/customer/match-or-create.ts`
- 입력: phone, name, source ('consultation'|'as'|'manual'), extra(주소 등)
- 동작: phone_normalized 검색 → 매칭됨 기존 customerId 반환 / 매칭 X 신규 INSERT
- 동명이인: 가장 오래된 customer로 매칭 (deterministic)
- phone 비어있으면 NULL 반환 (호출 측에서 NULL 처리)

**4 라우트에 helper 호출 추가**:
- `/api/consultation/public/submit` (source='consultation')
- `/api/consultation/admin-create` (source='manual')
- `/api/repair/public/submit` (source='as')
- `/api/repair` POST (source='manual', customer_id 없을 때만)

**백필 마이그 072** (선택, 사장님 SQL Editor에서 1번 실행):
- consultations / repairs / offline_sales의 NULL customer_id를 phone 매칭으로 1회성 채움

### 흐름
```
신규 phone 접수 → customers 신규 INSERT + customer_id 연결
같은 phone 재접수 → 기존 customer_id 매칭 (단골 자동 추적)
phone 없음 → customer_id NULL (예: 톡상담)
```

### 효과 — 판매 흐름 통합
- 상담 → "판매로 처리" CTA → /sales/new에서 SelectedCustomer 자동 set (customer_type 따라 가격 자동)
- 같은 고객 customer 상세 → 상담/복원수리/판매 모두 한 곳에 표시 (단골 식별)

### 회귀 안전
- 기존 동작 0 변경 — public/submit/admin-create는 customer_id가 채워질 뿐
- /sales/new from_consultation prefill 그대로 동작 (이미 customer_id 처리)
- workin 직접 판매 / customer-autocomplete / 070 link 모두 영향 0

---

## 2026-04-30 (밤) — 상담 리뷰 연동 전면 제거 (IA 단순화)

### 결정 배경
사장님 운영 현실: 매장방문 후 가위 안 사고 복원수리만 하거나 상담만 받고 가는 등 다양한 경우. 상담에서 후기요청 영역을 만들면 "어떤 후기?"라는 IA 모호함 발생. 한 거래(상담→가위/복원수리)는 거래 record(sale/repair)에서만 후기 관리하는 게 SSOT 친화.

### 제거된 영역
- **consultation-detail-panel**의 ReviewManagementCard 미노출 (매장/출장/톡상담 모두)
- **자동 후기요청 발송**: 상담완료 → review_request 알림톡 자동 트리거 제거 (`auto_request_on_completion` 토글도 무효)
- **약속 대기 탭** (`/reviews`)에서 consultation row 제거
- **B 칩** (related-activity) consultation 분기 제거
- **070 mirror 모드** (sale의 ReviewManagementCard) 제거 — sale은 일반 카드만

### 유지된 영역 (통계/이력 가치)
- DB 컬럼 `consultations.review_promised_at` / `review_request_sent_at` / `review_submitted_at` (NULL로 남거나 auto-match로 자동 채움)
- `/api/reviews/auto-match` (reviews/submit 시 source_id 매칭으로 자동 기록)

### 후기 흐름 — sale 단일 진입점
```
출장/매장상담 → "판매로 처리" CTA → /sales/new prefill (070 link) → sale 저장
  ↓ (sale 상세에서)
ReviewManagementCard → "후기 요청 보내기" 모달
  → review_type 자유 선택 (상담/복원수리/제품구매)
  → 알림톡 발송 → 후기 페이지 정상 동작
```

### 매장상담 → 복원수리만 한 케이스
- TMS에서 매장방문 상담 등록 (그대로)
- 별도 `/repairs/new`에서 복원수리 record 만들기 (고객 정보 직접 입력)
- 복원수리 record 카드에서 후기요청 (review_type='repair' 자동 고정)
- 향후 072 (상담→복원수리 link 인프라)로 입력 부담 절감 가능 (별도 task)

### 양방향 가시화
- consultation 상세에 linkedSales 칩 (녹색 [OS-...]) — 어떤 판매로 이어졌는지
- sale 상세에 linkedConsultation 정보 칩 (파란 [CS-...]) — 어떤 상담에서 시작됐는지

### 회귀 안전
- repair 모듈의 ReviewManagementCard 그대로 (별도 진입점)
- 070 link 인프라 (CTA / prefill / Modal / PATCH) 모두 유지
- consultation의 다른 흐름 (cancel/suggest/confirm/캘린더 sync) 0 영향
- DB 컬럼 보존 (auto-match 동작)

---

## 2026-04-30 (오전) — 출장요청 좌표 회귀 fix + Kakao geocode helper 통합

### 문제
고객 폼(`/api/consultation/public/submit`)에서 `geocodeAddress`가 address.json만 호출 + `addressRoad+addressDetail` 합친 query → 빌라/오피스텔 매칭 실패율 매우 높음 → 좌표 NULL → 지도 핀 미표시. admin-create는 이전에 fix됐지만 public/submit은 누락 (부분 fix 회귀).

### Fix
**신규 helper**: `app/src/lib/kakao/geocode.ts` — 3단계 fallback 표준화 (DRY)
1. address.json(road) — 도로명/지번 정확 매칭
2. keyword.json(road) — 신축/오래된 주소 유연 매칭
3. keyword.json(road+detail) — 상호명·건물명 케이스

**3 라우트 모두 helper 사용**: public/submit, admin-create, backfill-coords (이전엔 inline 함수 복제)

**클라이언트 안전망**: field-request-map.tsx 마운트 시 좌표 NULL 활성 출장요청 발견 → 자동으로 backfill-coords 호출 → GAS sync 등으로 들어온 NULL 건도 사장님이 지도 한 번 열면 자동 복구.

### 회귀 안전
- 기존 admin-create 외부 동작 0 변경 (helper로 추출만)
- field-request-map 필터 조건 그대로 (좌표 NULL 제외)

---

## 2026-04-30 (저녁) — 24h 리마인더 등록-방문 간격 가드

### 문제
당일 예약(예: 오늘 9시 등록 + 오늘 16시 visit) 케이스에 24h 리마인더가 발송됨. 의도: 24h 리마인더는 "방문 하루 전 안내"라 등록-방문 간격이 24h 미만이면 의미 없음.

### Fix — `send-reminders/route.ts`
24h 윈도우 통과 조건에 `intervalHours = (visit_at - created_at) / 1h >= 24` 가드 추가:
```ts
const intervalHours = (visitAt.getTime() - new Date(c.created_at).getTime()) / (1000 * 60 * 60);
if (diffHours <= 24 && diffHours > 2 && !c.remind_24h_at && intervalHours >= 24) { ... }
```

select 컬럼에 `created_at` 추가 (조회 보강).

**효과**:
- 당일 등록+visit (간격 < 24h) → 24h 리마인더 skip
- 어제 등록+오늘 visit (간격 ≥ 24h) → 정상 발송
- 2h 리마인더는 변화 0

---

## 2026-04-30 (저녁) — 후기요청 모달 IA 고정 (review_type 잠금 + subtype 자동)

### 문제
매장방문 후기요청 모달에서 사장님이 "복원수리" 옵션을 선택할 수 있어 → 솔라피 복원수리 템플릿(자체 정적 URL `?uid=#{as_uid}&type=as`)으로 발송 → 매장방문 UUID가 `type=as`로 후기 페이지 호출 → /api/reviews/info 'as' 미처리 → 404 오류.

### Fix — `review-request-modal.tsx`
**source 컨텍스트 기반 review_type 잠금**:

| source | review_type | UI |
|--------|------------|-----|
| consultation | `'consult'` 고정 | typeOptions 미노출, "자동 선택됨" 라벨 |
| repair | `'repair'` 고정 | 동일 |
| sale | `hasRepairItem`에 따라 default | typeOptions 노출 (사장님 변경 가능) |

**subtype 자동 추론**:
- 새 prop `sourceType?: string | null` 추가
- consultation → `consultations.consultation_type` 전달
- repair → `repairs.proceed_type` 전달
- 모달 초기값: `subtype = sourceType || 'store_visit'` (회귀 fallback)

### prop 전달 경로
```
consultation-detail-panel → ReviewManagementCard sourceType={c.consultation_type} → ReviewRequestModal
repair-detail-panel → ReviewManagementCard sourceType={r.proceed_type} → ReviewRequestModal
sale-detail-panel → ReviewManagementCard (sourceType 미전달) → ReviewRequestModal (sale은 그대로)
```

### 방어망 — `reviews/info`, `reviews/submit`에 type='as' alias
솔라피 측 정적 URL이 `?type=as`로 박혀있는 케이스 방어:
```ts
const type = rawType === 'as' ? 'repair' : rawType;
```

근본 fix는 모달 IA 잠금. alias는 안전망.

### 회귀 안전
- sale source 동작 0 변경 (사장님 typeOptions 선택 그대로)
- 자동 발송 흐름 (consultation/repair completed) 변화 0
- 070 mirror 모드 영향 0

---

## 2026-04-30 (오후) — 출장요청 지도 핀 IA 개선

### 문제
- 처리완료된 출장상담이 지도에 회색 `?` 핀으로 표시됨
- 원인: PIN_CONFIG에 `completed` 상태 누락 → DEFAULT_PIN fallback
- IA 관점: 활성 작업 탭에서 완료 핀이 떠 있는 건 시각 노이즈

### Fix — 컨텍스트별 핀 표시

**1. PIN_CONFIG에 `completed` 명시 추가** (`field-request-map.tsx`)
- 초록 ✓ — "성공/완료" 보편 매핑, on_hold(회색)와 차별화

**2. 필터 강화 — 컨텍스트별 동작**
| 탭 (activeStatuses) | 핀 표시 정책 |
|------|---------|
| 신규접수 / 제안중 / 일정재요청 / 확정 / 오늘출장 | 활성 status만, completed/cancelled 미노출 |
| **지난내역 (past)** | **completed 핀만 초록 ✓로 강조 표시** |
| 모든 탭 | cancelled은 항상 미노출 (위치 의미 X) |

**3. SUB_TAB_STATUSES 매핑 추가** (`consultations/page.tsx`)
- `past: ['completed']` 신규 추가 → 지난내역 탭이 활성일 때 completed 핀 트리거

### 핀 컬러 매트릭스 (전체)
| Status | 색상 | 아이콘 | 의미 |
|--------|------|-------|------|
| pending_admin | 노랑 #EAB308 | ! | 신규 접수 |
| suggested / reschedule / change_requested | 주황 #F97316 | →/↻ | 제안/재요청 |
| confirmed | 파랑 #3B82F6 | ✓ | 확정 |
| on_hold | 회색 #9CA3AF | ⏸ | 보류 |
| completed | 초록 #10B981 | ✓ | 처리완료 (지난내역에서만) |
| cancelled | — | — | 핀 X |

### 회귀 안전
- 활성 탭 동작: 완료 핀이 사라짐 → 사장님 IA 개선 의도
- 다른 모듈 (repair/sales/talk): 지도 X → 영향 0
- field-request-list 탭별 list 동작: 변경 0

---

## 2026-04-30 — GAS 의존 폐기 + 리마인더 회귀 fix

### GAS 트리거 OFF 완료
사장님 Apps Script 'MAMORU_Consulting' 프로젝트의 트리거 2개 모두 삭제:
- `sendReminders_` (15분 cron, 시트 기반 24h/2h 리마인더) — **삭제**
- `cleanupExpiredHolds_` (1시간 cron, 시트+캘린더 HOLD 정리) — **삭제**

GAS 코드(consulting/Code.gs, supabase-sync.gs, Total_Management_System/gas/consultation-sync.gs)는 1주 모니터링 후 archive 폴더로 이동 예정(완전 삭제 X).

### 리마인더 흐름 — TMS 단일 운영
```
Vercel Cron (10분 간격) → GET /api/cron/send-reminders
  → consultations WHERE status='confirmed' AND visit_date BETWEEN today AND tomorrow
  → 24h 윈도우 (2~24h 전) + remind_24h_at IS NULL → 마킹 후 sendNotification
  → 2h 윈도우 (0.5~2h 전) + remind_2h_at IS NULL → 마킹 후 sendNotification
  → Make webhook → Solapi 알림톡
```

### 리마인더 템플릿 매핑 (Make 시나리오 필터와 정확히 일치)
| consultation_type | 24h 템플릿 | 2h 템플릿 |
|------------------|----------|---------|
| `field_request` (출장) | `field_remind_24h` | `field_remind_2h` |
| `store_visit` (매장) | `remind24` | `remind2` |

⚠️ **회귀 fix 이력**: 2026-04-30 이전 코드는 매장/출장 모두 `field_remind_24h` / `field_remind_2h` 호출 → 매장 고객 메시지 발송 불일치 가능성. enum에 `remind24` / `remind2` 추가 + 분기 fix.

### 마이그 071 — remind_24h_at / remind_2h_at 컬럼 추가
TMS cron이 이미 사용 중이었던 컬럼이 DB에 누락 상태였음 → 마이그 071로 추가:
```sql
ALTER TABLE consultations
  ADD COLUMN IF NOT EXISTS remind_24h_at TIMESTAMPTZ NULL,
  ADD COLUMN IF NOT EXISTS remind_2h_at  TIMESTAMPTZ NULL;
CREATE INDEX idx_consult_reminder_pending
  ON consultations (visit_date, visit_time)
  WHERE status = 'confirmed';
```

### 폐기된 흐름 (참고용)
~~Google Sheets '상담접수' → onFormSubmit → /api/consultation/sync → Supabase~~
- 이미 onFormSubmit 트리거 없었음 (사장님 스크린샷 확인)
- 모든 신규 접수는 `/api/consultation/public/submit`으로 직접 들어옴
- consultation-sync.gs 파일은 archive 예정 (코드 보존)

상세: `memory/feedback_gas_deprecated.md` 참조

---

## 리뷰 요청 분기 (2026-04-29 추가)

상담완료(`completed`) 진입 시 후기 알림톡 발송 정책. `system_settings.review.auto_request_on_completion` 토글로 두 모드 양립.

### Mode A — 핀셋 정책 (default, OFF)
```
상담완료 클릭 → status='completed' → (자동 발송 X)
  → 사장님 판단:
       (고객 약속받음) ☑ 리뷰 약속 토글 → review_promised_at 기록
       (적당한 시점) [후기 요청 보내기] 클릭 → 알림톡 발송 + review_request_sent_at 기록
  → 고객 작성 → reviews/submit → consultations.review_submitted_at 자동 기록
  → 상세 패널 카드 = "작성 완료 ✓ · YYYY-MM-DD" 정적 라벨 / 약속 대기 탭에서 자동 사라짐
```

### Mode B — 안내문 정책 (toggle ON)
```
상담완료 클릭 → status='completed'
  → 자동 발송 가드 체크:
       review_promised_at IS NULL AND review_request_sent_at IS NULL → 자동 발송 + review_request_sent_at 기록
       review_promised_at !== NULL → 자동 skip (사장님 통제권)
       review_request_sent_at !== NULL → 자동 skip (중복 방지)
  → 약속 ✓ 고객은 항상 사장님 수동만
```

### 3 timestamp 의미
- `review_promised_at`: 사장님이 약속 토글 ✓ 한 시점
- `review_request_sent_at`: 후기 요청 알림톡 발송 시점 (자동/수동 무관)
- `review_submitted_at`: 고객이 리뷰 작성 완료 시점 (reviews/submit 역방향 매칭)

---

## 일정 수동 등록 모달 개선 (04-27)

`CreateConsultationModal`을 판매 입력 화면의 고객 입력 패턴과 일관되게 정렬:
- 고객명/연락처 텍스트 입력 → **`<CustomerAutocomplete>`** + 외부 "+ 신규 고객 등록"(중첩 `<CustomerCreateModal>`)
- 출장요청 주소 입력 → **`<DaumPostcodeButton>`** (postcode + road + detail 분리 저장)
- 출장요청일 때 `consultations.postcode` INSERT (컬럼 기존 존재, 마이그레이션 불필요)
- 매장↔출장 토글 시 `addressRoad`가 비어있을 때만 selectedCustomer에서 자동 채움 (사용자 임의 주소 보호)
- **자식 모달 떠 있는 동안 부모 dialog `open=false`**로 일시 닫기 → native `<dialog>` × 일반 fixed-div 중첩 시 가려짐 회피. 자식 닫히면 다시 열림. 입력값 손실 없음.
- ESC 가드: `showNewCustomer === true`면 부모 close 차단

**재사용 컴포넌트 추출**: `<DaumPostcodeButton>` 신설 + 기존 3곳(고객 신규/고객 수정/배송 설정)의 동일 코드 75줄 정리. CustomerAutocomplete에 `disableInlineNewForm` prop 추가(기본값 false → 다른 사용처 회귀 0).

## 일정변경 페이지 재진입 가드 (04-27)

`page_change_request.html`이 cancelled/completed/reschedule_requested 상태에서도 일정변경 폼이 그대로 노출되던 버그 — GAS getReservationInfo의 CONFIRMED/ASSIGNED 가드가 TMS Vercel API 이행 시 누락됐던 것.

**서버**: `GET /api/consultation/public/reservation`이 `canRequestChange: boolean` 반환 (confirmed/assigned만 true). backward compatible.

**페이지**: `loadReservation()`에서 `canRequestChange === false`면 status별 안내 화면 분기:
- `cancelled` → 기존 `cancel-done` 재사용 ("예약이 취소되었습니다")
- `completed` → 신규 `already-completed` ("이미 완료된 상담입니다")
- `reschedule_requested` → 신규 `already-rescheduled` ("이미 일정 변경 요청 중")
- 그 외(suggested 등) → 기존 `error` 화면

**카카오 인앱 닫기 fix** 동반: `safeClose`가 `kakaoweb://closeBrowser`(iOS 전용)만 호출하고 `return`으로 떨어져 Android에서 무반응이던 문제 → page_suggest/page_result 동일 패턴으로 UA 분기 + history.back→window.close fallback. mamoru.kr 강제 이동 폴백 제거(메모리 규칙 위반).

## cancel API 캘린더 동기화 (04-27)

admin-create / resched는 호출하던 `syncConsultationToCalendar`를 cancel API만 빠뜨려 DB는 cancelled인데 Google Calendar 일정이 잔류하던 문제. `after()` 블록으로 추가. `syncConsultationToCalendar`는 cancelled 상태에서 자동으로 `deleteCalendarEvent` + `google_event_id NULL` 처리 (기존 구현).

**잔여 정리 admin endpoint**: `GET /api/consultation/admin-cleanup-calendar` — cancelled + google_event_id IS NOT NULL 모든 행에 일괄 sync 호출. idempotent(두 번 호출 안전). 인증 필수.

## 푸시 알림 중복 도착 dedup (04-27)

리뷰/접수 등 푸시가 한 번 발송됐는데 사장님 디바이스에 두 번 표시되던 버그 — 두 가지 원인 동시 차단:
- **SW × onMessage 중복**: `firebase-messaging-sw.js`의 `push` 핸들러가 백그라운드/포그라운드 모두에서 `showNotification`을 호출하는데, `firebase/client.ts`의 `onMessage`가 또 `new Notification()` 생성 → 포그라운드에서 두 번 표시. → `onMessage` 핸들러 자체 제거(SW가 모든 표시 담당).
- **토큰 누적**: `push_subscriptions`에 한 사용자의 다중 토큰이 쌓여 모든 토큰에 발송되던 문제. → `subscribe API`에 single-token-per-user 정책(같은 user_id의 옛 토큰 자동 삭제). 신규 `cleanup-others` endpoint + 설정 → 알림 → "이 기기만 알림 받기" 버튼.

**알림 클릭 시 새 창 중복 차단**: SW의 `notificationclick` matchAll 매칭 범위를 same-origin 전체로 확장(기존: /dashboard·/consultations·/repairs 3경로만 → 다른 페이지에서 클릭 시 새 창 열렸음). focused 우선 → 기존 인스턴스 navigate + focus.

## 관리자 푸시 알림 (04-20 추가)

고객 행동 시 사장님 디바이스로 FCM 푸시 발송:
- **상담 접수 (매장방문)**: `confirmed` 템플릿 (접수 즉시 확정)
- **상담 접수 (출장)**: `field_request` 템플릿 (관리자 시간 제안 대기)
- **상담 접수 (톡상담)**: `talk_received` 템플릿
- **출장 일정 확정 ✅**: 고객이 제안된 시간 중 선택 (/api/consultation/public/confirm)
- **출장 일정 재요청 🔄**: 고객이 다른 시간 요청 (/api/consultation/public/resched)

설정: TMS → 설정 → 알림·연동 → "📱 내 푸시 알림" 섹션에서 개별 토글.

## 버그 수정 (04-20)
- **톡상담 취소 시 `cancelled` 템플릿(매장용) 발송 차단** — 톡상담은 알림톡 미발송 (카톡으로 이미 커뮤니케이션)

---

## 1. 비즈니스 프로세스 흐름

### 매장방문 (store_visit)
```
고객 접수 (아임웹 폼 → page_form.html)
  → Vercel API: Supabase 직접 저장 + 확정 알림톡 + Gmail 알림
  → [status: confirmed] (즉시 확정)
  → (고객) 셀프서비스 변경/취소 가능 (page_change_request.html → Vercel API)
  → [리마인드] D-1(24H) + D-0(2H) 자동 알림톡 (Vercel Cron 10분)
  → 방문 완료 → [status: completed]
  → 리뷰 요청 알림톡 → 리뷰 작성 (page_review.html)
```

### 출장요청 (field_request)
```
고객 접수 (아임웹 폼 → page_form.html)
  → Vercel API: Supabase 직접 저장 + 접수 알림톡(request) + Gmail 알림
  → [status: pending_admin]
  → (관리자 TMS) 시간 제안 (최대 3슬롯) → [status: suggested] + 제안 알림톡
  → (고객) 시간 선택 (page_suggest.html → Vercel API) → [status: confirmed] + 확정 알림톡
       또는 다른 일정 요청 → [status: reschedule_requested] + 관리자 이메일
  → [리마인드] D-1(24H) + D-0(2H) 자동 알림톡
  → (관리자) 출장 지연 시 → 지연안내 알림톡 (Vercel 직접)
  → 출장 완료 → [status: completed]
  → 리뷰 요청 알림톡
```

### 톡상담 (talk_consult)
```
고객 접수 (아임웹 폼 → page_form.html)
  → Vercel API: Supabase 직접 저장 + 접수 알림톡(talk_received)
  → [status: pending_admin]
  → (관리자 TMS) "상담 시작" → [status: in_progress] + talk_ready 알림톡
  → 상담 종료 → [status: completed]
  → 리뷰 요청 알림톡
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

## 2. 시스템 연동 흐름 (GAS 제거 완료)

```
[고객]
  │
  ▼
[아임웹 폼 / page_form.html]
  │ POST (JSON)
  ▼
[Vercel API — /api/consultation/public/submit]
  │
  ├──→ [Supabase DB] (직접 INSERT — ~20ms)
  ├──→ [Make Webhook → Solapi 알림톡]
  └──→ [Gmail 알림 (nodemailer)]

[TMS React UI]
  │ 관리자 액션
  ▼
[Vercel API — /api/consultation/*]
  │
  ├──→ [Supabase DB] (상태 변경 + 이력)
  └──→ [Make Webhook → Solapi 알림톡]

[고객 셀프서비스]
  │
  ├──→ [page_change_request.html → /api/consultation/public/cancel, /resched]
  ├──→ [page_suggest.html → /api/consultation/public/suggest-data, /confirm, /resched]
  └──→ [page_review.html → /api/reviews/submit]

[Vercel Cron — 10분 간격]
  └──→ [/api/cron/send-reminders → 24h/2h 리마인더 알림톡]
```

### 연동별 역할
| 시스템 | 역할 | 비고 |
|--------|------|------|
| Vercel API | 접수, 슬롯 조회, 알림톡, Gmail, 상태관리 | GAS 완전 대체 |
| Make Webhook | 알림톡 발송 허브 (17종 분기) | 기존과 동일 |
| Solapi | 카카오 알림톡 실제 발송 | 기존과 동일 |
| Supabase | 모든 데이터 저장 (SSOT) | Sheet 동기화 불필요 |
| Vercel Cron | 리마인더 24h/2h | GAS 트리거 대체 |
| **Google Calendar** | **상담 확정/변경/취소 미러링** | 2026-04-21 재도입 (OAuth 2.0) |
| ~~GAS~~ | ~~제거 완료~~ | 아카이브 보관 |
| ~~Google Sheets~~ | ~~제거 완료~~ | Supabase로 통합 |

---

## 3. 구현 완료 ✅

### 공개 API Routes (고객 접수/셀프서비스 — 비인증)
| 엔드포인트 | 메서드 | 기능 |
|------------|--------|------|
| `/api/consultation/public/submit` | POST | 고객 접수 (매장/출장/톡) |
| `/api/consultation/public/settings` | GET | 영업시간 + 휴무 설정 |
| `/api/consultation/public/slots` | GET | 예약 가능 슬롯 조회 (~50ms) |
| `/api/consultation/public/reservation` | GET | 예약 정보 조회 (uid) |
| `/api/consultation/public/cancel` | GET | 고객 취소 요청 |
| `/api/consultation/public/confirm` | GET | 고객 시간 선택 확정 |
| `/api/consultation/public/resched` | GET | 고객 일정 재요청 |
| `/api/consultation/public/suggest-data` | GET | 시간 제안 데이터 조회 |

### TMS API Routes (관리자 — 인증 필수)
| 엔드포인트 | 메서드 | 기능 |
|------------|--------|------|
| `/api/consultation` | GET/POST | 목록/생성 (기본 CRUD) |
| `/api/consultation/admin-create` | POST | **수기 등록 전용** (인스타DM/유선 접수 → 확정 + 알림톡 + 캘린더 + 중복체크) |
| `/api/consultation/[id]` | GET/PATCH | 상세/상태변경 + 알림톡 |
| `/api/consultation/suggest` | POST | 시간 제안 + 알림톡 (직접) |
| `/api/consultation/delay` | POST | 출장 지연 안내 + 알림톡 (직접) |
| `/api/consultation/notify` | POST | 수동 알림톡 발송 |
| `/api/cron/send-reminders` | GET | 24h/2h 리마인더 (Cron) |

### 알림톡 (17종 — Make webhook 경유)
```
매장방문 (5종): confirmed, cancelled, rescheduled, remind24, remind2
출장요청 (9종): request, suggest, field_confirmed, field_cancelled,
                field_rescheduled, field_remind_24h, field_remind_2h,
                field_delayed, change_request_received
톡상담   (2종): talk_received, talk_ready
공통     (1종): review_request (상담완료 후 리뷰)
```

### 고객 대면 페이지 (GitHub Pages)
| 페이지 | 용도 | API |
|--------|------|-----|
| page_form.html | 상담 접수 | Vercel public/submit + slots + settings |
| page_suggest.html | 출장 시간 선택 | Vercel public/suggest-data + confirm + resched |
| page_change_request.html | 예약 변경/취소 | Vercel public/reservation + cancel + resched |
| page_review.html | 리뷰 작성 | Vercel reviews/submit |
| page_diag.html | 간편 진단 | 클라이언트 전용 (API 없음) |
| page_guide.html | 서비스 안내 | 정적 페이지 |
| page_recommend.html | 제품 추천 | 클라이언트 전용 |

---

## 4. GAS → Vercel 이전 히스토리 (2026-04-01 완료)

| 항목 | 완료일 |
|------|--------|
| DB: consultation_settings + closed_dates 테이블 | 03-31 |
| API: 접수/슬롯/설정 공개 엔드포인트 | 03-31 |
| API: 시간제안 GAS→직접 알림톡 | 04-01 |
| API: 출장지연 GAS→직접 알림톡 | 04-01 |
| API: 취소 GAS cancelViaGAS 제거 | 04-01 |
| 고객 페이지 JS 교체 (GAS→Vercel) | 03-31~04-01 |
| 셀프서비스 API 5종 신규 | 04-01 |
| Gmail 발송 nodemailer 구현 | 04-01 |
| Vercel Pro 업그레이드 + Cron 10분 | 04-01 |
| Google Calendar 제거 | 04-01 |

---

## 5. Google Calendar 연동 (2026-04-21 재도입)

### 목적
TMS 상담을 사장님의 `bsm@mamoru.kr` Google Calendar에 자동 미러링 → **모바일 기본 달력 위젯에서 일정 즉시 확인**.
※ SSOT(진실의 원천)은 여전히 TMS DB. 캘린더는 "표시용" 거울.

### 상태별 캘린더 액션 매트릭스
| 상담 상태 | 캘린더 표시 | 이벤트 액션 | 제목 예시 |
|----------|------------|------------|----------|
| `pending_admin` / `assigned` | ❌ | — | — |
| `suggested` | ❌ | — (고객 미확정) | — |
| `confirmed` | ✅ 생성/업데이트 | CREATE or UPDATE | `[매장] 홍길동 · 010-xxxx` |
| `reschedule_requested` / `change_requested` | ⚠️ 유지 | UPDATE (⏳ 노란색) | `[출장] ⏳ 김철수 · 강남구 · 010-xxxx` |
| `on_hold` | ❌ | DELETE | — |
| `completed` | ✅ 유지 | UPDATE (✅ 회색) | `[매장] ✅ 박민수 · 010-xxxx` |
| `cancelled` | ❌ | DELETE | — |

### 트리거 지점 (모두 Next.js `after()` 래퍼 안에서 실행 — Vercel 서버리스 보장)
| 경로 | 트리거 | 동작 |
|------|--------|------|
| `POST /api/consultation/public/submit` | 매장방문 즉시 confirmed | CREATE |
| `GET /api/consultation/public/confirm` | 출장 고객 시간 수락 | CREATE |
| `GET /api/consultation/public/resched` | 고객 재요청 | UPDATE (⏳) |
| `PATCH /api/consultation/[id]` | 관리자 상태/시간 변경 | CREATE/UPDATE/DELETE 자동 분기 |
| `DELETE /api/consultation/[id]` | 관리자 삭제 | DELETE |

### 이벤트 필드
- **제목**: `[매장/출장] (⏳/✅)? 이름 · 지역(출장만) · 연락처`
- **위치(Location)**: 매장 주소(매장방문) 또는 고객 주소(출장) — 탭 시 지도 앱 자동 실행
- **시간**: visit_date + visit_time, 종료 = 시작 + 60분(기본) — KST
- **색상**: 매장 파랑(1) / 출장 녹색(2) / 재요청 노랑(5) / 완료 회색(8)
- **설명**: 상담종류·고객명·연락처·주소·메모·상담번호·접수일시·TMS 링크·tel: 링크
- **알림**: 기본 OFF (알림톡·푸시 중복 방지)
- **숨김 메타**: `mamoru_consultation_id` 등 extendedProperties.private

### 파일 구조
```
app/src/lib/google/
  ├ oauth.ts              # OAuth2 클라이언트 + 토큰 관리 (system_settings)
  ├ calendar-client.ts    # events insert/patch/delete wrapper
  ├ event-formatter.ts    # Consultation → Event 변환
  └ calendar-sync.ts      # 상태별 분기 orchestrator

app/src/app/api/google/calendar/
  ├ auth/         # GET: OAuth 인가 URL
  ├ callback/     # GET: code 교환 + 토큰 저장 + Workspace 자동 판별
  ├ disconnect/   # POST: 토큰 삭제
  ├ status/       # GET: 연결 상태 조회
  └ resync/       # POST: 과거 60일~미래 180일 일괄 동기화

app/src/components/settings/
  └ google-calendar-settings.tsx  # 설정 UI (알림·연동 탭 최상단)
```

### 환경변수 (Vercel)
```
GOOGLE_CLIENT_ID          # OAuth 2.0 웹 클라이언트 ID
GOOGLE_CLIENT_SECRET      # OAuth 2.0 클라이언트 시크릿 (Sensitive)
GOOGLE_REDIRECT_URI       # https://app-eta-sandy-75.vercel.app/api/google/calendar/callback
```

### OAuth 스코프
- `openid email profile` — id_token 발급 (연결 계정 email/hd 자동 판별)
- `https://www.googleapis.com/auth/calendar.events` — 이벤트 CRUD

### 주의사항
- **재요청 시 매장방문 슬롯은 해제됨** (기존 로직 유지) — 캘린더는 이벤트 유지하되 제목에 ⏳ 표기
- **캘린더에서 직접 수정 금지** — SSOT 위반 방지 (description에 경고 문구)
- **refresh_token 6개월 무사용 시 자동 무효** → 설정 UI에 경고 배너 + 재연결 필요

---

## 6. 관리자 직접 상담 등록 — 수기 접수 흐름 (2026-04-24 추가)

### 목적
인스타 DM · 유선 전화 · 매장 워크인 등 **외부 채널로 들어온 상담**을 사장님이 TMS에서 직접 입력하여 기존 자동화 흐름(알림톡·리마인더·캘린더·푸시)에 자연스럽게 편입시킴.

### 지원 범위
| 유형 | 지원 | 비고 |
|------|------|------|
| 매장방문 (store_visit) | ✅ | 주소 불필요 |
| 출장요청 (field_request) | ✅ | 주소 필수 (지도 표시 + 리마인더 문자) |
| 톡상담 (talk_consult) | ❌ | 일정 없는 유형 → 수기 등록 제외 |

### 진입점
- **상담관리 페이지 → 탭 행 우측 "일정수동등록" 버튼**
- 통합 모달에서 유형 선택 → 필드 동적 노출

### 입력 필드
| 필드 | 매장방문 | 출장요청 | 비고 |
|------|:-:|:-:|------|
| 유형 (세그먼트) | ✅ | ✅ | 기본 매장방문 |
| 고객명 | ✅ | ✅ | 중복 검사 기준 아님 |
| 연락처 | ✅ | ✅ | **중복 검사 기준** (phone_normalized) |
| 방문 날짜 | ✅ | ✅ | 필수 |
| 방문 시간 | ✅ | ✅ | 필수 |
| 주소 | ❌ | ✅ | 지오코딩으로 lat/lng 자동 설정 |
| 상세 주소 | ❌ | (선택) | 상호명/층수 등 |
| 메모 | (선택) | (선택) | 접수 경로 기록 권장 |
| 알림톡 발송 체크 | ✅ | ✅ | 기본 On |

### 자동 처리 흐름 (등록 버튼 클릭 시)
```
POST /api/consultation/admin-create
├─ 관리자 인증 체크
├─ 입력 검증 (타입별 필수 분기)
├─ 중복 체크
│   └─ phone_normalized + visit_date + visit_time + 미래 → 409 Conflict
│       → 경고 모달에서 기존 상담 정보 카드 표시
├─ 출장 지오코딩 (Kakao REST API)
├─ DB INSERT
│   ├─ status: 'confirmed' (즉시 확정)
│   ├─ unique_id: crypto.randomUUID()
│   └─ gas_raw.source: 'admin_manual' (수기 마킹)
├─ consultation_history INSERT ('관리자 직접 등록')
└─ after() {
     (notify=true) → sendNotification('confirmed' / 'field_confirmed')
     → syncConsultationToCalendar() (Google Calendar 이벤트 자동 생성)
   }
```

### 자동 편입되는 기존 흐름
| 기능 | 자동 편입 여부 | 근거 |
|------|:-:|------|
| 24h/2h 리마인더 | ✅ | cron이 status=confirmed + visit_date/time 조회 |
| Google Calendar 동기화 | ✅ | calendar-sync가 확정 상태 감지 |
| 상태 변경 (취소/완료/일정변경) | ✅ | 기존 상세 패널 액션 전부 적용 |
| 고객 변경/취소 링크 | ✅ | change_request_link 알림톡에 포함 |
| 리포트 집계 | ✅ | `gas_raw.source = 'admin_manual'` 필터로 수기 N건 / 웹 M건 구분 |

### 중복 감지 정책
- **기준**: `phone_normalized` (하이픈 유무 무관) + visit_date + visit_time
- **범위**: 오늘 이후 미래 건만 (과거 같은 고객은 무시)
- **활성 상태**: pending_admin / assigned / suggested / confirmed / reschedule_requested / change_requested
- **충돌 시**: 409 Conflict + `existing` 상담 객체 반환 → 경고 모달에서 "기존 확인" or "입력 폼 복귀"

### 파일 구조
```
app/src/app/api/consultation/admin-create/route.ts  # 신규 API
app/src/hooks/use-consultations.ts                  # useCreateConsultation 훅
app/src/components/consultations/
  └ create-consultation-modal.tsx                   # 신규 모달 + 중복 경고 서브 모달
app/src/app/(dashboard)/consultations/page.tsx      # 버튼 + 모달 연결
```

### 설계 원칙
- **단일 진실점**: 수기 접수와 웹 접수가 **동일 consultations 테이블 + 동일 상태머신** 사용
- **유지보수 단순화**: 기존 알림톡/리마인더/캘린더 로직을 전부 재사용 (분기 최소화)
- **수기 구분 가능**: `gas_raw.source` 필드로 리포트에서만 구분 (운영 흐름은 동일)
- **UX 일관성**: 기존 상세 패널, 필터, 검색 모두 수기 등록 건에도 적용
