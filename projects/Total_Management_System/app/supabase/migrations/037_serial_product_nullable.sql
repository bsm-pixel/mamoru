-- 이카운트 이관 시리얼 등 product_id 없이도 저장 가능하도록
ALTER TABLE product_serials ALTER COLUMN product_id DROP NOT NULL;
