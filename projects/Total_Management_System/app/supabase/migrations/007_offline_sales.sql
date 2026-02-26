-- 007: 오프라인 판매 테이블 + 이카운트 연동 필드
-- R5: 오프라인 판매 + 이카운트 ERP 연동

-- 1) customers 테이블에 이카운트 거래처 코드 추가
ALTER TABLE customers ADD COLUMN IF NOT EXISTS ecount_customer_code text DEFAULT NULL;

-- 2) 오프라인 판매 테이블
CREATE TABLE IF NOT EXISTS offline_sales (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sale_number text NOT NULL UNIQUE,                -- OS-20260226-001 형식
  customer_id uuid REFERENCES customers(id),
  customer_name text NOT NULL,
  customer_phone text,
  sale_date date NOT NULL DEFAULT CURRENT_DATE,
  total_amount integer NOT NULL DEFAULT 0,
  discount_amount integer NOT NULL DEFAULT 0,
  paid_amount integer NOT NULL DEFAULT 0,
  payment_method text NOT NULL DEFAULT 'card',     -- card | cash | transfer | mixed
  payment_status text NOT NULL DEFAULT 'paid',     -- paid | unpaid | partial
  memo text,
  ecount_sync_status text NOT NULL DEFAULT 'pending', -- pending | synced | failed
  ecount_slip_no text,                             -- 이카운트 전표번호
  ecount_synced_at timestamptz,
  created_by uuid REFERENCES profiles(id),
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

-- 3) 오프라인 판매 항목 테이블
CREATE TABLE IF NOT EXISTS offline_sale_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sale_id uuid NOT NULL REFERENCES offline_sales(id) ON DELETE CASCADE,
  product_id uuid REFERENCES products(id),
  product_name text NOT NULL,
  sku text,
  quantity integer NOT NULL DEFAULT 1,
  unit_price integer NOT NULL DEFAULT 0,
  total_price integer NOT NULL DEFAULT 0
);

-- 4) 인덱스
CREATE INDEX IF NOT EXISTS idx_offline_sales_date ON offline_sales(sale_date DESC);
CREATE INDEX IF NOT EXISTS idx_offline_sales_customer ON offline_sales(customer_id);
CREATE INDEX IF NOT EXISTS idx_offline_sales_ecount ON offline_sales(ecount_sync_status);
CREATE INDEX IF NOT EXISTS idx_offline_sale_items_sale ON offline_sale_items(sale_id);

-- 5) updated_at 트리거
CREATE OR REPLACE FUNCTION update_offline_sales_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_offline_sales_updated_at ON offline_sales;
CREATE TRIGGER trg_offline_sales_updated_at
  BEFORE UPDATE ON offline_sales
  FOR EACH ROW
  EXECUTE FUNCTION update_offline_sales_updated_at();

-- 6) RLS (기본 정책: 인증 사용자 전체 접근)
ALTER TABLE offline_sales ENABLE ROW LEVEL SECURITY;
ALTER TABLE offline_sale_items ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can manage offline_sales"
  ON offline_sales FOR ALL TO authenticated USING (true) WITH CHECK (true);

CREATE POLICY "Authenticated users can manage offline_sale_items"
  ON offline_sale_items FOR ALL TO authenticated USING (true) WITH CHECK (true);
