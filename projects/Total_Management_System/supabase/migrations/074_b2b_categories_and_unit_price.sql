-- 074: B2B 카테고리 동적 관리 + customer_product_catalog에 unit_price 추가
-- 사장님이 설정에서 B2B 카테고리(딜러/아카데미/학교/공기관 등) 자유롭게 추가 가능
-- 또한 거래처별 맞춤 단가 catalog에 등록 가능
-- Supabase SQL Editor에서 실행

-- 1) system_settings에 b2b.categories 기본값 seed (이미 있으면 무시)
INSERT INTO system_settings (key, value)
VALUES (
  'b2b.categories',
  '[
    {"key": "dealer",  "label": "딜러",     "icon": "Users",         "display_order": 1, "is_active": true, "is_default": true},
    {"key": "academy", "label": "아카데미", "icon": "GraduationCap", "display_order": 2, "is_active": true, "is_default": true}
  ]'::jsonb
)
ON CONFLICT (key) DO NOTHING;

-- 2) customer_product_catalog에 unit_price 컬럼 추가 (거래처별 맞춤 가격)
ALTER TABLE customer_product_catalog
  ADD COLUMN IF NOT EXISTS unit_price integer NULL;

COMMENT ON COLUMN customer_product_catalog.unit_price
  IS '거래처별 맞춤 가격. NULL이면 product.price_groups[customerType].price 또는 product.price fallback.';
