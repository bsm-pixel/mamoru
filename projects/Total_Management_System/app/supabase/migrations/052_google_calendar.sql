-- 052_google_calendar.sql — Google Calendar 연동
-- 상담 확정 시 자동으로 Google Calendar에 이벤트 생성/업데이트/삭제
-- bsm@mamoru.kr 계정에 OAuth 연결 후 동작

-- 1. consultations에 구글 이벤트 매칭 컬럼 추가
ALTER TABLE consultations
  ADD COLUMN IF NOT EXISTS google_event_id TEXT,
  ADD COLUMN IF NOT EXISTS google_event_updated_at TIMESTAMPTZ;

CREATE INDEX IF NOT EXISTS idx_consultations_google_event_id
  ON consultations(google_event_id) WHERE google_event_id IS NOT NULL;

COMMENT ON COLUMN consultations.google_event_id IS 'Google Calendar 이벤트 ID (동기화된 건만)';
COMMENT ON COLUMN consultations.google_event_updated_at IS 'Google Calendar 이벤트 마지막 동기화 시각';

-- 2. OAuth 토큰/상태는 system_settings에 저장 (마이그레이션 불필요 — 런타임에 UPSERT)
-- system_settings에 저장될 키:
--   google.calendar.access_token         — 수명 1h
--   google.calendar.refresh_token        — 장기 유지
--   google.calendar.token_expires_at     — ISO 8601
--   google.calendar.connected_email      — 연결된 Google 계정
--   google.calendar.connected_hd         — Workspace 도메인 (일반 Gmail이면 빈값)
--   google.calendar.connected_at         — 최초 연결 시각
--   google.calendar.calendar_id          — 기본 'primary'
--   google.calendar.last_error           — 최근 실패 메시지
--   google.calendar.last_success_at      — 최근 성공 시각

-- 3. 매장 정보는 system_settings (business.store_address는 기존 존재, store_name 신규)
--   business.store_address   — 매장방문 이벤트 Location (기존)
--   business.store_name      — 매장 이름 (신규)
