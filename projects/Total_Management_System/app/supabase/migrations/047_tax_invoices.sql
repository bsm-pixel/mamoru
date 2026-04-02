-- 047: 세금계산서 발행 관리
CREATE TABLE IF NOT EXISTS tax_invoices (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  invoice_type text NOT NULL CHECK (invoice_type IN ('sales', 'purchase')),
  issue_date date NOT NULL DEFAULT CURRENT_DATE,
  counterparty_name text NOT NULL,
  counterparty_biz_no text,
  supply_amount integer NOT NULL DEFAULT 0,
  tax_amount integer NOT NULL DEFAULT 0,
  total_amount integer NOT NULL DEFAULT 0,
  sale_id uuid,
  purchase_order_id uuid,
  memo text,
  created_by uuid,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_tax_inv_date ON tax_invoices(issue_date);
CREATE INDEX IF NOT EXISTS idx_tax_inv_type ON tax_invoices(invoice_type);
CREATE INDEX IF NOT EXISTS idx_tax_inv_sale ON tax_invoices(sale_id);
