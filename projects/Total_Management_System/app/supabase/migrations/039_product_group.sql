-- 039_product_group.sql — 리뷰 제품군 그룹핑 (product_group)
-- 동일 시리즈 제품(R4-58ST, R4-58-MS 등)의 리뷰를 묶어서 표시하기 위한 컬럼

ALTER TABLE products ADD COLUMN IF NOT EXISTS product_group TEXT;
ALTER TABLE reviews ADD COLUMN IF NOT EXISTS product_group TEXT;

CREATE INDEX IF NOT EXISTS idx_products_product_group ON products(product_group);
CREATE INDEX IF NOT EXISTS idx_reviews_product_group ON reviews(product_group);

COMMENT ON COLUMN products.product_group IS '리뷰 그룹핑 키 (예: R4, M5, CS600) — 동일 시리즈 제품은 같은 값';
COMMENT ON COLUMN reviews.product_group IS '리뷰 제출 시 제품에서 복사된 그룹 키';
