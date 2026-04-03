-- 048: 판매 송장 + 고객 우편번호
ALTER TABLE customers ADD COLUMN IF NOT EXISTS postcode text;

ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS invoice_number text;
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS shipped_at timestamptz;
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS courier_name text DEFAULT '롯데택배';
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS delivery_method text DEFAULT 'pickup';
-- delivery_method: 'pickup'(직접수령) | 'shipping'(택배발송)
