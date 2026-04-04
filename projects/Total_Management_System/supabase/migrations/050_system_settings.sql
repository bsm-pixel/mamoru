-- 050: system_settings 테이블 (TMS 설정 탭 리뉴얼)
-- Supabase SQL Editor에서 직접 실행

CREATE TABLE IF NOT EXISTS system_settings (
  key TEXT PRIMARY KEY,
  value JSONB NOT NULL DEFAULT '{}',
  updated_at TIMESTAMPTZ DEFAULT now()
);

ALTER TABLE system_settings ENABLE ROW LEVEL SECURITY;

CREATE POLICY "auth_read" ON system_settings FOR SELECT USING (auth.role() = 'authenticated');
CREATE POLICY "auth_write" ON system_settings FOR ALL USING (auth.role() = 'authenticated');

-- 반품 기능용 컬럼
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS returned_at TIMESTAMPTZ;
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS return_reason TEXT;

-- 다중 창고
CREATE TABLE IF NOT EXISTS warehouses (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL,
  description TEXT,
  sort_order INT DEFAULT 0,
  created_at TIMESTAMPTZ DEFAULT now()
);

ALTER TABLE warehouses ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_warehouses" ON warehouses FOR ALL USING (auth.role() = 'authenticated');

INSERT INTO warehouses (name, description, sort_order) VALUES
  ('보관창고', '입고 후 대기', 1),
  ('준비창고', '판매 대기 (시리얼 부여)', 2)
ON CONFLICT DO NOTHING;

-- 기본 설정값 삽입 (현재 하드코딩 값을 그대로 기본값으로)
INSERT INTO system_settings (key, value) VALUES
  -- 대시보드
  ('dashboard.monthly_goal', '0'),
  ('dashboard.low_stock_threshold', '3'),
  ('dashboard.repair_stale_days', '3'),
  ('dashboard.kpi_green', '80'),
  ('dashboard.kpi_yellow', '50'),
  ('dashboard.outstanding_warning', '0'),
  ('dashboard.monthly_purchase_limit', '0'),
  ('dashboard.card_visibility', '{"sales":true,"repairs":true,"orders":true,"consultations":true,"outstanding":true,"lowStock":true}'),
  ('dashboard.card_order', '["sales","repairs","orders","consultations","outstanding","lowStock"]'),
  -- 복원수리
  ('repair.price_mamoru', '10000'),
  ('repair.price_other', '20000'),
  ('repair.shipping_fees', '[{"qty":1,"fee":5000},{"qty":2,"fee":3000},{"qty":3,"fee":0}]'),
  ('repair.extra_services', '[{"name":"날 변형 (곡률 조절)","price":50000}]'),
  ('repair.bank_account', '{"bank":"","number":"","holder":""}'),
  ('repair.stale_days', '3'),
  ('repair.estimated_days', '7'),
  ('repair.unpaid_reminder_days', '3'),
  ('repair.inspection_categories', '{"blade_tip":["양호","무뎌짐","찍힘"],"blade_mid":["양호","무뎌짐","찍힘"],"blade_inner":["양호","무뎌짐","찍힘"],"comb":["양호","손상","벌어짐"],"tension":["양호","헐거움","과다"],"parts":["양호","이상"],"stopper":["양호","이상"]}'),
  -- 배송
  ('shipping.sender', '{"name":"마모루","tel":"","zip":"","addr":""}'),
  ('shipping.default_goods_name', '"가위 복원수리"'),
  ('shipping.memo_presets', '["부재 시 경비실","파손주의 - 가위","배송 전 연락 부탁"]'),
  ('shipping.unshipped_warning_days', '2'),
  ('shipping.review_delay_days', '0'),
  ('shipping.auto_push_invoice', 'false'),
  -- 알림
  ('notifications.master_enabled', 'true'),
  ('notifications.consultation_received', 'true'),
  ('notifications.repair_received', 'true'),
  ('notifications.repair_cost_notice', 'true'),
  ('notifications.repair_payment_confirmed', 'true'),
  ('notifications.repair_shipped', 'true'),
  ('notifications.review_request', 'true'),
  ('notifications.sound_enabled', 'false'),
  ('notifications.sound_targets', '{"orders":true,"consultations":true,"repairs":true}'),
  -- 판매
  ('sales.payment_methods', '["card","cash","transfer","mixed"]'),
  ('sales.channels', '["offline","online","talk"]'),
  ('sales.receipt_footer', '""'),
  ('sales.dealer_discount', '0'),
  ('sales.academy_discount', '0'),
  ('sales.block_zero_stock', 'false'),
  ('sales.default_vat_included', 'true'),
  ('sales.auto_tax_invoice', 'false'),
  -- 고객
  ('customer.types', '["retail","online","dealer","academy"]'),
  ('customer.sources', '["imweb","consultation","as","manual"]'),
  ('customer.rfm', '{"recency_vip":90,"recency_dormant":180,"frequency":3,"monetary":1000000}'),
  ('customer.outstanding_reminder_days', '0'),
  ('customer.default_sort', '"name"'),
  ('customer.memo_templates', '[]'),
  ('customer.tags', '[]'),
  -- 상품·재고
  ('inventory.low_stock_threshold', '3'),
  ('inventory.imweb_sync', 'true'),
  ('inventory.categories', '["BL","TH","LO","SL","CB","CS","AC"]'),
  ('inventory.safety_stock', '5'),
  ('inventory.adjustment_reasons', '["파손","분실","증정","샘플","실사 조정","기타"]'),
  ('inventory.sku_digits', '3'),
  ('inventory.serial_auto_trigger', '"manual"'),
  ('inventory.stocktake_reminder_day', '0'),
  ('inventory.barcode_format', '"Code128"'),
  ('inventory.default_sort', '"name"'),
  -- 회계
  ('accounting.vat_rate', '10'),
  ('accounting.expense_categories', '["택배비","포장재","교통비","사무용품","식대","소모품","임대료","인건비","기타"]'),
  ('accounting.bank_accounts', '[]'),
  ('accounting.tax_type', '"general"'),
  ('accounting.revenue_basis', '"sale_date"'),
  ('accounting.fiscal_start_month', '1'),
  ('accounting.budgets', '{}'),
  -- 상담
  ('consultation.reminder_24h_enabled', 'true'),
  ('consultation.reminder_2h_enabled', 'true'),
  ('consultation.auto_review_request', 'false'),
  ('consultation.gmail_notify', 'true'),
  ('consultation.duration_by_type', '{"store_visit":60,"field_request":90,"talk_consult":0}'),
  ('consultation.change_deadline_hours', '0'),
  -- 시스템
  ('system.table_page_size', '20'),
  ('system.start_page', '"/dashboard"'),
  ('system.sidebar_config', '{"order":[],"hidden":[]}'),
  -- 사업자 정보 (거래명세서+세금계산서+시스템 공유)
  ('business.info', '{"company":"","registration_number":"","representative":"","address":"","phone":"","business_type":"","business_item":""}'),
  ('business.logo_url', '""'),
  ('business.store_address', '""'),
  ('business.store_lat', '0'),
  ('business.store_lng', '0')
ON CONFLICT (key) DO NOTHING;