-- 093: 복원수리 직접방문(당일수리) Google Calendar 연동 — consultations(052) 패턴 동일
--
-- 배경 (2026-05-25 Phase 3-B):
--   사장님 비전 — 복원수리 직접방문 일정이 Google Calendar 에도 자동 박힘
--   기존 컨설팅 매장방문/출장 Google Calendar 동기화 인프라 그대로 재사용:
--     · lib/google/calendar-client.ts (createEvent/updateEvent/deleteEvent)
--     · lib/google/oauth.ts            (getAuthorizedClient — refresh_token 자동 갱신)
--     · lib/google/calendar-sync.ts    (syncConsultationToCalendar — 직접방문용 syncRepairToCalendar 신규 작성)
--     · lib/google/event-formatter.ts  (formatConsultationToEvent — formatRepairToEvent 추가)
--   사장님이 이미 OAuth 인증 완료 + 컨설팅 동기화 작동 중이므로 환경변수/토큰 추가 작업 X
--
-- 변경 (consultations 052 와 동일 패턴):
--   repairs 테이블에 2컬럼 추가:
--     - google_event_id          (TEXT, Google Calendar 이벤트 ID)
--     - google_event_updated_at  (TIMESTAMPTZ, 마지막 동기화 시각)
--
-- 영향:
--   기존 데이터 무영향 (nullable). system_settings 변경 X (기존 키 그대로).

ALTER TABLE repairs
  ADD COLUMN IF NOT EXISTS google_event_id          TEXT,
  ADD COLUMN IF NOT EXISTS google_event_updated_at  TIMESTAMPTZ;

COMMENT ON COLUMN repairs.google_event_id         IS 'Google Calendar 이벤트 ID (직접방문 동기화용, NULL=미동기화)';
COMMENT ON COLUMN repairs.google_event_updated_at IS 'Google Calendar 마지막 동기화 시각';

-- partial index (NULL 제외, 동기화된 건만 빠른 조회)
CREATE INDEX IF NOT EXISTS idx_repairs_google_event_id
  ON repairs(google_event_id)
  WHERE google_event_id IS NOT NULL;

-- 검증 SQL (사장님이 Supabase SQL Editor 에서 직접 확인)
-- SELECT column_name, data_type FROM information_schema.columns
-- WHERE table_name='repairs' AND column_name LIKE 'google_%';
-- → 2행: google_event_id(text), google_event_updated_at(timestamp with time zone)
