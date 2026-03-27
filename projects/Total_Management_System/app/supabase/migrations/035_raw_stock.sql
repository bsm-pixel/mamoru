-- 보관창고 수량 (비시리얼, 매입 원본)
-- stock_quantity = raw_stock + COUNT(in_stock 시리얼)
ALTER TABLE products ADD COLUMN IF NOT EXISTS raw_stock integer NOT NULL DEFAULT 0;

-- 기존 stock_quantity가 시리얼 없이 수량만 있던 제품은 raw_stock으로 이관
-- (시리얼이 없는 제품의 stock_quantity를 raw_stock으로 복사)
UPDATE products p
SET raw_stock = CASE
  WHEN p.stock_quantity > 0 AND NOT EXISTS (
    SELECT 1 FROM product_serials ps WHERE ps.product_id = p.id AND ps.status = 'in_stock'
  ) THEN p.stock_quantity
  ELSE 0
END;
