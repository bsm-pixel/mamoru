-- 061: 매입품목 카탈로그 + 부가세 유형 확장
-- Supabase SQL Editor에서 실행

-- 1. 매입품목 카탈로그 테이블
CREATE TABLE IF NOT EXISTS supplier_product_catalog (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  supplier_id uuid NOT NULL REFERENCES customers(id) ON DELETE CASCADE,
  product_id uuid NOT NULL REFERENCES products(id) ON DELETE CASCADE,
  order_name text,       -- 주문명 (공장 발주용, 품명과 다를 수 있음)
  features text,         -- 특징/메모
  created_at timestamptz DEFAULT now(),
  UNIQUE(supplier_id, product_id)
);
CREATE INDEX IF NOT EXISTS idx_spc_supplier ON supplier_product_catalog(supplier_id);

-- 2. 발주서 부가세 유형 (기존 is_vat_included boolean 유지 + 새 컬럼)
ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS vat_type text DEFAULT 'included';
-- 값: 'included'(포함) / 'separate'(별도) / 'none'(미적용)

-- 3. 매입처별 기본 부가세 유형
ALTER TABLE customers ADD COLUMN IF NOT EXISTS default_vat_type text DEFAULT 'included';
