-- 020: 재고 수동 조정 이력 테이블
CREATE TABLE IF NOT EXISTS stock_adjustments (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  product_id uuid NOT NULL REFERENCES products(id),
  adjustment_type text NOT NULL CHECK (adjustment_type IN ('damage', 'correction', 'return', 'other')),
  quantity integer NOT NULL,  -- 양수=증가, 음수=감소
  reason text,
  adjusted_by uuid REFERENCES profiles(id),
  created_at timestamptz DEFAULT now()
);

CREATE INDEX idx_stock_adj_product ON stock_adjustments(product_id);
CREATE INDEX idx_stock_adj_created ON stock_adjustments(created_at DESC);

-- RLS
ALTER TABLE stock_adjustments ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_users_all" ON stock_adjustments FOR ALL TO authenticated USING (true) WITH CHECK (true);
