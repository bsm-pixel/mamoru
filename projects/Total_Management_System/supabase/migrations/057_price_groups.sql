-- 057: 동적 단가 그룹 시스템
-- products 테이블에 price_groups JSONB 컬럼 추가
-- 기존 price_dealer/price_academy/dealer_name/academy_name 데이터를 마이그레이션

-- 1) JSONB 컬럼 추가
ALTER TABLE products
  ADD COLUMN IF NOT EXISTS price_groups jsonb DEFAULT '{}'::jsonb;

-- 2) 기존 데이터 마이그레이션
UPDATE products
SET price_groups = jsonb_strip_nulls(jsonb_build_object(
  'dealer', jsonb_strip_nulls(jsonb_build_object(
    'price', CASE WHEN price_dealer > 0 THEN price_dealer ELSE null END,
    'display_name', CASE WHEN dealer_name IS NOT NULL AND dealer_name != '' THEN dealer_name ELSE null END
  )),
  'academy', jsonb_strip_nulls(jsonb_build_object(
    'price', CASE WHEN price_academy > 0 THEN price_academy ELSE null END,
    'display_name', CASE WHEN academy_name IS NOT NULL AND academy_name != '' THEN academy_name ELSE null END
  ))
))
WHERE price_dealer > 0
   OR price_academy > 0
   OR (dealer_name IS NOT NULL AND dealer_name != '')
   OR (academy_name IS NOT NULL AND academy_name != '');

-- 3) settings에 단가 그룹 정의 시드
INSERT INTO system_settings (key, value)
VALUES (
  'pricing.groups',
  '{"dealer":{"label":"딜러가","color":"purple","customerTypes":["dealer"]},"academy":{"label":"아카데미가","color":"emerald","customerTypes":["academy"]}}'
)
ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;

-- 기존 컬럼(price_dealer, price_academy, dealer_name, academy_name)은 전환기 동안 유지
