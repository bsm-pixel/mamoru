-- 판매 취소 시 원래 zone으로 복원하기 위한 컬럼
ALTER TABLE product_serials ADD COLUMN IF NOT EXISTS previous_zone text;
