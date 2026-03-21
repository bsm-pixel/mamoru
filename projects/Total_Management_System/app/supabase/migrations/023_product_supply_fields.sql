-- 023: 부자재 관리용 필드
ALTER TABLE products ADD COLUMN IF NOT EXISTS purchase_url text;          -- 주문 링크 (네이버/알리 등)
ALTER TABLE products ADD COLUMN IF NOT EXISTS supply_status text DEFAULT 'sufficient';  -- sufficient/needed/ordered
