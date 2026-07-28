-- 121: 복원수리 고객 셀프 관리 토큰 (일정확인/변경·취소 링크용 — as_id 순차노출 대신 랜덤 토큰)
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS manage_token uuid DEFAULT gen_random_uuid();
-- 기존 행 백필 (DEFAULT는 신규행만 적용되므로)
UPDATE repairs SET manage_token = gen_random_uuid() WHERE manage_token IS NULL;
ALTER TABLE repairs ALTER COLUMN manage_token SET NOT NULL;
CREATE UNIQUE INDEX IF NOT EXISTS idx_repairs_manage_token ON repairs(manage_token);
COMMENT ON COLUMN repairs.manage_token IS '고객 셀프서비스(일정확인/변경/취소) 링크 토큰. page_change_request.html?uid=<token> (121, 2026-07-28)';
