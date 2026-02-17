-- ============================================
-- MAMORU TMS Phase 2-2A — 상담 리뉴얼 DB 확장
-- ============================================

-- ENUM 확장: consultation_type에 톡상담 추가
ALTER TYPE consultation_type ADD VALUE IF NOT EXISTS 'talk_consult';

-- ENUM 확장: consultation_status에 보류/진행중/처리완료 추가
ALTER TYPE consultation_status ADD VALUE IF NOT EXISTS 'on_hold';
ALTER TYPE consultation_status ADD VALUE IF NOT EXISTS 'in_progress';
ALTER TYPE consultation_status ADD VALUE IF NOT EXISTS 'completed';

-- consultations 컬럼 추가
ALTER TABLE consultations ADD COLUMN IF NOT EXISTS hold_reason TEXT;
ALTER TABLE consultations ADD COLUMN IF NOT EXISTS latitude DOUBLE PRECISION;
ALTER TABLE consultations ADD COLUMN IF NOT EXISTS longitude DOUBLE PRECISION;

-- 인덱스 추가
CREATE INDEX IF NOT EXISTS idx_consultations_type_status ON consultations(consultation_type, status);
CREATE INDEX IF NOT EXISTS idx_consultations_coords ON consultations(latitude, longitude) WHERE latitude IS NOT NULL;
CREATE INDEX IF NOT EXISTS idx_consultations_visit_date ON consultations(visit_date) WHERE visit_date IS NOT NULL;
