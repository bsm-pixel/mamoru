-- 015: 매입(발주) 관리 테이블
CREATE TABLE purchase_orders (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  po_number text NOT NULL UNIQUE,
  supplier_id uuid REFERENCES customers(id),
  supplier_name text NOT NULL,
  order_date date NOT NULL DEFAULT CURRENT_DATE,
  expected_date date,
  received_date date,
  total_amount bigint NOT NULL DEFAULT 0,
  deposit_amount bigint DEFAULT 0,
  deposit_paid_at timestamptz,
  balance_amount bigint DEFAULT 0,
  balance_paid_at timestamptz,
  is_vat_included boolean DEFAULT true,
  supply_amount bigint DEFAULT 0,
  vat_amount bigint DEFAULT 0,
  status text NOT NULL DEFAULT 'draft',
  memo text,
  created_by uuid REFERENCES profiles(id),
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

CREATE TABLE purchase_order_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  po_id uuid NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE,
  product_id uuid REFERENCES products(id),
  product_name text NOT NULL,
  sku text,
  quantity integer NOT NULL DEFAULT 1,
  unit_price bigint NOT NULL DEFAULT 0,
  total_price bigint NOT NULL DEFAULT 0
);

CREATE INDEX idx_po_status ON purchase_orders(status);
CREATE INDEX idx_po_supplier ON purchase_orders(supplier_id);
CREATE INDEX idx_poi_po ON purchase_order_items(po_id);
