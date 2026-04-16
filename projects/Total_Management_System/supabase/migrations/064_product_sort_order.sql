-- 제품 진열 순서: product_group → sort_order → name 정렬용
ALTER TABLE products ADD COLUMN IF NOT EXISTS sort_order integer DEFAULT 0;
