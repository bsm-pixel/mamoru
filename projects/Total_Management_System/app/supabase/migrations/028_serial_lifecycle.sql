-- 028: 시리얼 Lifecycle 강화 — warehouse_zone 3단계 (raw/ready/display)

-- 1) 기존 storage → raw 변환, NULL도 raw로
UPDATE product_serials SET warehouse_zone = 'raw'
  WHERE warehouse_zone = 'storage' OR warehouse_zone IS NULL;

-- 2) 기본값 변경
ALTER TABLE product_serials ALTER COLUMN warehouse_zone SET DEFAULT 'raw';
-- warehouse_zone 값: 'raw' (매입원본/마킹전) | 'ready' (마킹+포장완료) | 'display' (쇼케이스)
