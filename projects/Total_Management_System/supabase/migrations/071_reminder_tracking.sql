-- 071: consultations.remind_24h_at / remind_2h_at 컬럼 추가
-- TMS의 /api/cron/send-reminders가 이미 사용 중인데 마이그 누락 상태였음
-- 중복 발송 방지용 timestamp — null이면 미발송, 있으면 발송 완료 시각 기록
-- Supabase SQL Editor에서 실행

ALTER TABLE consultations
  ADD COLUMN IF NOT EXISTS remind_24h_at TIMESTAMPTZ NULL,
  ADD COLUMN IF NOT EXISTS remind_2h_at  TIMESTAMPTZ NULL;

COMMENT ON COLUMN consultations.remind_24h_at IS '24h 리마인더 발송 시각 — 중복 발송 방지 (null=미발송)';
COMMENT ON COLUMN consultations.remind_2h_at  IS '2h 리마인더 발송 시각 — 중복 발송 방지 (null=미발송)';

-- send-reminders cron 조회 성능 개선 (confirmed 상태 + 미래 일정만 스캔)
CREATE INDEX IF NOT EXISTS idx_consult_reminder_pending
  ON consultations (visit_date, visit_time)
  WHERE status = 'confirmed';
