# 상담관리 프로세스 흐름도
> 최종 업데이트: 2026-04-20 (관리자 푸시 알림 확장 + 톡상담 취소 알림톡 버그 수정)

---

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
| ~~GAS~~ | ~~제거 완료~~ | 아카이브 보관 |
| ~~Google Calendar~~ | ~~제거 완료~~ | TMS 달력으로 대체 |
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
| `/api/consultation` | GET/POST | 목록/생성 |
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
