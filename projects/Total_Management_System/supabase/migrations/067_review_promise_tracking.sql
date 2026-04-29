-- 067: 리뷰 약속·요청 추적 시스템 — 자동/수동 모드 통합
-- consultations / repairs / offline_sales에 timestamp 3종 추가
-- system_settings에 자동 발송 정책 토글 시드 (기본 false = 핀셋 동봉 정책 호환)
-- Supabase SQL Editor에서 실행

-- 상담관리: 약속·요청·작성 시점 추적
ALTER TABLE consultations
  ADD COLUMN IF NOT EXISTS review_promised_at      timestamptz,
  ADD COLUMN IF NOT EXISTS review_request_sent_at  timestamptz,
  ADD COLUMN IF NOT EXISTS review_submitted_at     timestamptz;

-- 복원수리: 동일
ALTER TABLE repairs
  ADD COLUMN IF NOT EXISTS review_promised_at      timestamptz,
  ADD COLUMN IF NOT EXISTS review_request_sent_at  timestamptz,
  ADD COLUMN IF NOT EXISTS review_submitted_at     timestamptz;

-- 오프라인 판매: review_requested_at 이미 존재 (rename 시 6 callsites 회귀) → 신규 2개만
ALTER TABLE offline_sales
  ADD COLUMN IF NOT EXISTS review_promised_at  timestamptz,
  ADD COLUMN IF NOT EXISTS review_submitted_at timestamptz;

COMMENT ON COLUMN offline_sales.review_requested_at
  IS 'semantic alias of review_request_sent_at (legacy column kept for backward compat)';

-- 약속 대기 조회용 partial index — review_promised_at != null AND review_submitted_at == null
CREATE INDEX IF NOT EXISTS idx_consult_review_pending
  ON consultations (review_promised_at)
  WHERE review_promised_at IS NOT NULL AND review_submitted_at IS NULL;

CREATE INDEX IF NOT EXISTS idx_repair_review_pending
  ON repairs (review_promised_at)
  WHERE review_promised_at IS NOT NULL AND review_submitted_at IS NULL;

CREATE INDEX IF NOT EXISTS idx_offsale_review_pending
  ON offline_sales (review_promised_at)
  WHERE review_promised_at IS NOT NULL AND review_submitted_at IS NULL;

-- 자동 발송 정책 토글 시드 — false = 현재 핀셋 동봉 정책 호환
INSERT INTO system_settings (key, value)
  VALUES ('review.auto_request_on_completion', 'false')
  ON CONFLICT (key) DO NOTHING;
