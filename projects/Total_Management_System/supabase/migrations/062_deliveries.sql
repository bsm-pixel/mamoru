-- 062: 납품관리 모듈 (B2B 딜러/아카데미 납품)
-- Supabase SQL Editor에서 실행

-- 1. 납품 마스터 테이블
CREATE TABLE IF NOT EXISTS deliveries (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  dl_number text NOT NULL UNIQUE,
  customer_id uuid REFERENCES customers(id),
  customer_name text NOT NULL,
  customer_phone text,
  customer_type text,
  delivery_date date NOT NULL DEFAULT CURRENT_DATE,
  expected_date date,
  shipped_date date,
  total_amount bigint NOT NULL DEFAULT 0,
  discount_amount bigint DEFAULT 0,
  paid_amount bigint DEFAULT 0,
  payment_status text DEFAULT 'unpaid',
  payment_method text DEFAULT 'transfer',
  vat_type text DEFAULT 'included',
  supply_amount bigint DEFAULT 0,
  vat_amount bigint DEFAULT 0,
  receipt_type text DEFAULT 'expense_proof',
  tracking_number text,
  status text NOT NULL DEFAULT 'draft',
  memo text,
  cancelled_at timestamptz,
  cancelled_reason text,
  created_by uuid,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

-- 2. 납품 항목 테이블
CREATE TABLE IF NOT EXISTS delivery_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  delivery_id uuid NOT NULL REFERENCES deliveries(id) ON DELETE CASCADE,
  product_id uuid REFERENCES products(id),
  product_name text NOT NULL,
  sku text,
  category text,
  quantity integer NOT NULL DEFAULT 1,
  unit_price bigint NOT NULL DEFAULT 0,
  total_price bigint NOT NULL DEFAULT 0
);

-- 인덱스
CREATE INDEX IF NOT EXISTS idx_deliveries_status ON deliveries(status);
CREATE INDEX IF NOT EXISTS idx_deliveries_customer ON deliveries(customer_id);
CREATE INDEX IF NOT EXISTS idx_deliveries_date ON deliveries(delivery_date);
CREATE INDEX IF NOT EXISTS idx_delivery_items_delivery ON delivery_items(delivery_id);
