-- 014: 제품 3단 가격 + 매입처 + 아임웹 매핑 + 설명
ALTER TABLE products ADD COLUMN IF NOT EXISTS price_dealer bigint DEFAULT 0;     -- 딜러/도매가
ALTER TABLE products ADD COLUMN IF NOT EXISTS price_purchase bigint DEFAULT 0;   -- 매입가(원가)
ALTER TABLE products ADD COLUMN IF NOT EXISTS supplier_id uuid REFERENCES customers(id);
ALTER TABLE products ADD COLUMN IF NOT EXISTS description text;
ALTER TABLE products ADD COLUMN IF NOT EXISTS imweb_product_no text;             -- 아임웹 상품번호 매핑

CREATE INDEX IF NOT EXISTS idx_products_supplier ON products(supplier_id);
CREATE INDEX IF NOT EXISTS idx_products_imweb ON products(imweb_product_no);
