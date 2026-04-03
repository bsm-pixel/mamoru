-- 049: 시리얼 → 판매 항목 정확 매칭
ALTER TABLE product_serials ADD COLUMN IF NOT EXISTS sale_item_id uuid;
