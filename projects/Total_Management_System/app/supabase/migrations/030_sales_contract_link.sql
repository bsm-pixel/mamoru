-- 030: 판매 ↔ 계약서 양방향 연결
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS contract_id uuid REFERENCES contracts(id) DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_offline_sales_contract ON offline_sales(contract_id);
