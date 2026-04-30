-- 073: B2B 납품처별 납품품목 카탈로그 (mirror: supplier_product_catalog)
-- dealer/academy 별로 같은 product를 다른 이름으로 납품할 수 있게 — 송장에 그 이름이 박혀서 출력
-- supplier_product_catalog (매입처용)와 별개 테이블 (회귀 안전 + 의미 분리)
-- Supabase SQL Editor에서 실행

CREATE TABLE IF NOT EXISTS customer_product_catalog (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  customer_id uuid NOT NULL REFERENCES customers(id) ON DELETE CASCADE,
  product_id uuid NOT NULL REFERENCES products(id) ON DELETE CASCADE,
  delivery_name text,
  features text,
  sort_order integer DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE(customer_id, product_id)
);

CREATE INDEX IF NOT EXISTS idx_cpc_customer ON customer_product_catalog(customer_id);
CREATE INDEX IF NOT EXISTS idx_cpc_product ON customer_product_catalog(product_id);

COMMENT ON TABLE customer_product_catalog
  IS 'B2B 납품처(dealer/academy)별 납품품목 카탈로그. customer_type=supplier는 별도 supplier_product_catalog 사용.';
COMMENT ON COLUMN customer_product_catalog.delivery_name
  IS '송장/납품서/거래명세서/준비표에 출력될 납품명. 각 customer별 다르게 설정 가능. NULL이면 product.price_groups display_name 또는 product.name fallback.';
COMMENT ON COLUMN customer_product_catalog.features
  IS '규격/특이사항 — 발주서/명세서에 함께 출력 (옵션)';
