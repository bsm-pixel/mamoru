-- 009: 시리얼넘버/바코드 관리
-- R7: products 확장 + 개별 시리얼 추적

-- 1) 시리얼넘버 테이블 (개별 제품 추적)
CREATE TABLE IF NOT EXISTS product_serials (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  product_id uuid NOT NULL REFERENCES products(id),
  serial_number text NOT NULL UNIQUE,              -- MM-BL001-20260226-001 형식
  barcode text UNIQUE,                             -- 바코드 (EAN-13 or 자체 형식)
  status text NOT NULL DEFAULT 'in_stock',         -- in_stock | reserved | sold | returned | defective
  -- 판매 연결
  sold_via text,                                   -- online | offline | contract
  order_id uuid REFERENCES orders(id),
  offline_sale_id uuid REFERENCES offline_sales(id),
  contract_id uuid REFERENCES contracts(id),
  sold_at timestamptz,
  sold_to_name text,
  sold_to_phone text,
  -- 메타
  lot_number text,                                 -- 생산 로트
  manufactured_at date,                            -- 제조일
  memo text,
  created_by uuid REFERENCES profiles(id),
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

-- 2) 인덱스
CREATE INDEX IF NOT EXISTS idx_serials_product ON product_serials(product_id);
CREATE INDEX IF NOT EXISTS idx_serials_status ON product_serials(status);
CREATE INDEX IF NOT EXISTS idx_serials_barcode ON product_serials(barcode);
CREATE INDEX IF NOT EXISTS idx_serials_order ON product_serials(order_id);
CREATE INDEX IF NOT EXISTS idx_serials_sale ON product_serials(offline_sale_id);
CREATE INDEX IF NOT EXISTS idx_serials_contract ON product_serials(contract_id);

-- 3) updated_at 트리거
CREATE OR REPLACE FUNCTION update_product_serials_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_product_serials_updated_at ON product_serials;
CREATE TRIGGER trg_product_serials_updated_at
  BEFORE UPDATE ON product_serials
  FOR EACH ROW
  EXECUTE FUNCTION update_product_serials_updated_at();

-- 4) RLS
ALTER TABLE product_serials ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can manage product_serials"
  ON product_serials FOR ALL TO authenticated USING (true) WITH CHECK (true);
