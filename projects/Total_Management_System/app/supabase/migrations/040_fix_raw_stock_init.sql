-- 040_fix_raw_stock_init.sql — 기존 상품 raw_stock 보정
-- 아임웹에서 동기화된 상품 중 raw_stock=0이고 stock_quantity>0인 것을 보관창고에 채우기
-- 단, 이미 시리얼이 등록된 상품은 제외 (시리얼로 전환된 수량이 있으므로)

UPDATE products
SET raw_stock = stock_quantity
WHERE raw_stock = 0
  AND stock_quantity > 0
  AND is_active = true
  AND id NOT IN (
    SELECT DISTINCT product_id
    FROM product_serials
    WHERE product_id IS NOT NULL
      AND status = 'in_stock'
  );
