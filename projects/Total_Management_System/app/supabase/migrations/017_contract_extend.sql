-- 017_contract_extend.sql
-- 계약서 전자문서 확장: 수령방법, 선납/잔금, 판매자 서명, 매장 정보

ALTER TABLE contracts ADD COLUMN IF NOT EXISTS delivery_method text DEFAULT 'shipping';
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS unavailable_days text;
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS deposit_amount bigint DEFAULT 0;
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS balance_amount bigint DEFAULT 0;
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS seller_signature text;
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS customer_title text;
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS shop_name text;
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS shop_address text;
