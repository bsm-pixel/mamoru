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
