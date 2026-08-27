-- 134_returns_exchange_out.sql — 교환 출고(새 제품 발송) 추적 필드 (2026-08-27, 교환/반품 Phase 2.2)
-- 배송 교환 시 "새 제품(교환해줄 것)"을 롯데 송장으로 별도 출고. returns가 반품(구제품 회수)만 담고 있어
-- 새 제품 정보 + 교환 출고 송장(매장→고객, 정방향)을 담을 컬럼 추가.
-- 기존 invoice_number/courier_name(132)은 "수거 송장(고객→매장 역방향)"용이라 방향이 반대 → 별도 컬럼.

ALTER TABLE returns ADD COLUMN IF NOT EXISTS new_product_id            uuid;
ALTER TABLE returns ADD COLUMN IF NOT EXISTS new_product_name          text;
ALTER TABLE returns ADD COLUMN IF NOT EXISTS new_serial_number         text;
ALTER TABLE returns ADD COLUMN IF NOT EXISTS exchange_out_invoice_number text;
ALTER TABLE returns ADD COLUMN IF NOT EXISTS exchange_out_courier_name text;
ALTER TABLE returns ADD COLUMN IF NOT EXISTS exchange_shipped_at       timestamptz;

COMMENT ON COLUMN returns.new_product_name IS '교환해줄 새 제품명(교환 출고 송장 품목명)';
COMMENT ON COLUMN returns.exchange_out_invoice_number IS '교환 출고 송장번호(매장→고객, 새 제품 발송)';
