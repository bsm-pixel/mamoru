-- 006: 복원수리 접수확인/포장완료 타임스탬프 추가
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS confirmed_at timestamptz DEFAULT NULL;
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS packed_at    timestamptz DEFAULT NULL;

COMMENT ON COLUMN repairs.confirmed_at IS '관리자 접수확인 시점';
COMMENT ON COLUMN repairs.packed_at    IS '포장완료 시점 (출고 전 확인용)';
