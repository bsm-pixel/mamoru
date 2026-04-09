-- 058: offline_sale_items에 category 컬럼 추가
-- 복원수리(RS) 매출 집계를 product_name ILIKE 대신 category 기반으로 전환

-- 1) category 컬럼 추가
ALTER TABLE offline_sale_items
  ADD COLUMN IF NOT EXISTS category text;

-- 2) 기존 데이터 마이그레이션: product_id가 있는 항목은 products.category에서 복사
UPDATE offline_sale_items osi
SET category = p.category
FROM products p
WHERE osi.product_id = p.id
  AND osi.category IS NULL;

-- 3) product_name에 '복원수리' 포함된 항목은 RS로 설정 (product_id 없는 경우 포함)
UPDATE offline_sale_items
SET category = 'RS'
WHERE category IS NULL
  AND product_name ILIKE '%복원수리%';
