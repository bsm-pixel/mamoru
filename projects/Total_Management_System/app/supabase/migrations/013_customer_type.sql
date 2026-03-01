-- 013: 고객 유형 + 메모 + 미수금
ALTER TABLE customers ADD COLUMN IF NOT EXISTS customer_type text DEFAULT 'retail';
-- retail(일반) | online(아임웹) | dealer(딜러/도매) | supplier(매입처)

ALTER TABLE customers ADD COLUMN IF NOT EXISTS company_name text;
ALTER TABLE customers ADD COLUMN IF NOT EXISTS memo text;
ALTER TABLE customers ADD COLUMN IF NOT EXISTS outstanding_balance bigint DEFAULT 0;
