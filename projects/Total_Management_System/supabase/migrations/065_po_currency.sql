-- 발주 외화 지원: 통화 + 환율
ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS currency text DEFAULT 'KRW';
ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS exchange_rate numeric DEFAULT 1;
