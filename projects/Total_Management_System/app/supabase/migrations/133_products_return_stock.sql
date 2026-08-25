-- 133_products_return_stock.sql — 비시리얼 제품 반품창고 재고 (2026-08-25, 교환/반품 Phase 2.1)
-- 시리얼 제품은 product_serials(status=returned, zone=return)로 반품창고 추적하지만,
-- 비시리얼 제품(빗·악세서리·test 등)은 개별 추적번호가 없어 제품별 카운터로 관리.
-- 판매가능 현재고(stock_quantity)·아임웹과 분리(검수대기).

ALTER TABLE products ADD COLUMN IF NOT EXISTS return_stock integer NOT NULL DEFAULT 0;
COMMENT ON COLUMN products.return_stock IS '반품창고 재고(비시리얼) — 검수대기·판매가능 현재고/아임웹 미반영';
