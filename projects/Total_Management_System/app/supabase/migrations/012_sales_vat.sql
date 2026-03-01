-- 012: 판매 VAT 분리 컬럼 추가
-- 카드결제 시 공급가액/부가세 자동 계산하여 세금계산서 준비 자료로 활용

ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS supply_amount integer DEFAULT 0;
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS vat_amount integer DEFAULT 0;
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS is_vat_included boolean DEFAULT true;

ALTER TABLE offline_sale_items ADD COLUMN IF NOT EXISTS supply_amount integer DEFAULT 0;
ALTER TABLE offline_sale_items ADD COLUMN IF NOT EXISTS vat_amount integer DEFAULT 0;
