-- ============================================
-- MAMORU TMS Phase 2-1 — 상담/딜러 스키마
-- ============================================

-- ENUM 타입
CREATE TYPE consultation_status AS ENUM (
  'pending_admin', 'suggested', 'assigned', 'confirmed',
  'cancelled', 'reschedule_requested'
);

CREATE TYPE consultation_type AS ENUM ('store_visit', 'field_request');
CREATE TYPE dealer_status AS ENUM ('active', 'inactive');

-- ============================================
-- dealers: 딜러 정보
-- ============================================
CREATE TABLE dealers (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  dealer_code TEXT UNIQUE NOT NULL,
  name TEXT NOT NULL,
  phone TEXT,
  regions TEXT[] NOT NULL DEFAULT '{}',
  calendar_id TEXT,
  status dealer_status NOT NULL DEFAULT 'active',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_dealers_code ON dealers(dealer_code);
CREATE INDEX idx_dealers_status ON dealers(status);

ALTER TABLE dealers ENABLE ROW LEVEL SECURITY;

CREATE POLICY "dealers_all_authenticated" ON dealers
  FOR ALL USING (auth.role() = 'authenticated');

-- ============================================
-- consultations: 상담 접수
-- ============================================
CREATE TABLE consultations (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  customer_id UUID REFERENCES customers(id) ON DELETE SET NULL,
  name TEXT NOT NULL,
  phone TEXT NOT NULL,
  phone_normalized TEXT GENERATED ALWAYS AS (regexp_replace(phone, '\D', '', 'g')) STORED,
  consultation_type consultation_type NOT NULL DEFAULT 'store_visit',
  visit_date DATE,
  visit_time TEXT,
  postcode TEXT,
  address_road TEXT,
  address_detail TEXT,
  address_sido TEXT,
  address_sigungu TEXT,
  address_region TEXT,
  status consultation_status NOT NULL DEFAULT 'pending_admin',
  memo TEXT,
  unique_id TEXT UNIQUE NOT NULL,
  dealer_id UUID REFERENCES dealers(id) ON DELETE SET NULL,
  suggestions JSONB,
  gas_raw JSONB,
  received_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_consultations_status ON consultations(status);
CREATE INDEX idx_consultations_customer ON consultations(customer_id);
CREATE INDEX idx_consultations_dealer ON consultations(dealer_id);
CREATE INDEX idx_consultations_unique ON consultations(unique_id);
CREATE INDEX idx_consultations_received ON consultations(received_at DESC);
CREATE INDEX idx_consultations_phone ON consultations(phone_normalized);

ALTER TABLE consultations ENABLE ROW LEVEL SECURITY;

CREATE POLICY "consultations_all_authenticated" ON consultations
  FOR ALL USING (auth.role() = 'authenticated');

-- ============================================
-- consultation_history: 상담 상태 이력
-- ============================================
CREATE TABLE consultation_history (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  consultation_id UUID NOT NULL REFERENCES consultations(id) ON DELETE CASCADE,
  from_status consultation_status,
  to_status consultation_status NOT NULL,
  changed_by UUID REFERENCES profiles(id) ON DELETE SET NULL,
  note TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_consultation_history_consultation ON consultation_history(consultation_id);

ALTER TABLE consultation_history ENABLE ROW LEVEL SECURITY;

CREATE POLICY "consultation_history_all_authenticated" ON consultation_history
  FOR ALL USING (auth.role() = 'authenticated');

-- ============================================
-- updated_at 트리거 (기존 함수 재활용)
-- ============================================
CREATE TRIGGER trg_dealers_updated BEFORE UPDATE ON dealers
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER trg_consultations_updated BEFORE UPDATE ON consultations
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
