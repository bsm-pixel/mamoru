-- Phase 7: 복원수리 시스템 테이블
-- Supabase SQL Editor에서 실행

-- 1) repairs (메인)
CREATE TABLE IF NOT EXISTS repairs (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  as_id TEXT UNIQUE NOT NULL,              -- AS-YYYYMMDD-NNN (GAS 발번)
  customer_id UUID REFERENCES customers(id),
  name TEXT NOT NULL,
  phone TEXT NOT NULL,
  phone_normalized TEXT,
  proceed_type TEXT,                        -- 방문수거/카운터보관/직접전달/직접발송
  postcode TEXT,
  address TEXT,
  address_detail TEXT,
  pickup_date TEXT,
  delivery_method TEXT,
  qty_mamoru INT DEFAULT 0,
  qty_other INT DEFAULT 0,
  memo TEXT,
  service_cost INT DEFAULT 0,              -- 순수 수리비
  shipping_fee INT DEFAULT 0,              -- 수거비
  total_amount INT DEFAULT 0,              -- 합계
  status TEXT NOT NULL DEFAULT 'intake',
  invoice_number TEXT,                     -- 출고 송장
  courier_name TEXT DEFAULT '롯데택배',
  shipped_at TIMESTAMPTZ,
  delivered_at TIMESTAMPTZ,
  admin_note TEXT,                         -- 관리자 메모
  gas_raw JSONB,                           -- GAS 원본 데이터
  received_at TIMESTAMPTZ DEFAULT now(),
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

-- phone_normalized 자동 생성 트리거
CREATE OR REPLACE FUNCTION repair_normalize_phone()
RETURNS TRIGGER AS $$
BEGIN
  NEW.phone_normalized := regexp_replace(NEW.phone, '\D', '', 'g');
  NEW.updated_at := now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_repair_normalize_phone
  BEFORE INSERT OR UPDATE ON repairs
  FOR EACH ROW EXECUTE FUNCTION repair_normalize_phone();

-- 인덱스
CREATE INDEX IF NOT EXISTS idx_repairs_status ON repairs(status);
CREATE INDEX IF NOT EXISTS idx_repairs_phone_normalized ON repairs(phone_normalized);
CREATE INDEX IF NOT EXISTS idx_repairs_received_at ON repairs(received_at DESC);
CREATE INDEX IF NOT EXISTS idx_repairs_as_id ON repairs(as_id);

-- 2) repair_inspections (가위별 검수)
CREATE TABLE IF NOT EXISTS repair_inspections (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  repair_id UUID REFERENCES repairs(id) ON DELETE CASCADE NOT NULL,
  scissor_number INT NOT NULL,             -- 가위 번호 (1, 2, 3...)
  scissor_type TEXT,                       -- 블런트/틴닝/장가위/슬라이싱
  blade_tip TEXT DEFAULT '양호',
  blade_mid TEXT DEFAULT '양호',
  blade_inner TEXT DEFAULT '양호',
  comb TEXT DEFAULT '',                    -- 틴닝 전용
  tension TEXT DEFAULT '양호',
  parts TEXT DEFAULT '양호',               -- 내부부품
  stopper TEXT DEFAULT '양호',
  photo_url TEXT,
  photo_marks JSONB,                       -- [{x, y, label}]
  worker TEXT DEFAULT '백성민',
  created_at TIMESTAMPTZ DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_repair_inspections_repair_id ON repair_inspections(repair_id);

-- 3) repair_history (상태 이력)
CREATE TABLE IF NOT EXISTS repair_history (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  repair_id UUID REFERENCES repairs(id) ON DELETE CASCADE NOT NULL,
  from_status TEXT,
  to_status TEXT NOT NULL,
  changed_by UUID,
  note TEXT,
  created_at TIMESTAMPTZ DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_repair_history_repair_id ON repair_history(repair_id);

-- 4) RLS 정책
ALTER TABLE repairs ENABLE ROW LEVEL SECURITY;
ALTER TABLE repair_inspections ENABLE ROW LEVEL SECURITY;
ALTER TABLE repair_history ENABLE ROW LEVEL SECURITY;

-- 인증된 사용자 전체 접근 (TMS 관리자)
CREATE POLICY "repairs_auth_all" ON repairs
  FOR ALL USING (auth.role() = 'authenticated');

CREATE POLICY "repair_inspections_auth_all" ON repair_inspections
  FOR ALL USING (auth.role() = 'authenticated');

CREATE POLICY "repair_history_auth_all" ON repair_history
  FOR ALL USING (auth.role() = 'authenticated');

-- Service Role은 RLS 우회 (GAS sync용)
