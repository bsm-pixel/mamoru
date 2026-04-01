-- 043: 시리얼 previous_zone 컬럼 추가 (판매 취소 시 zone 복원용)
-- 기존 코드에서 previous_zone을 읽고/쓰는데 DB 컬럼이 없었음

ALTER TABLE product_serials
ADD COLUMN IF NOT EXISTS previous_zone text;

COMMENT ON COLUMN product_serials.previous_zone IS '판매 전 warehouse_zone (취소 시 복원용)';
