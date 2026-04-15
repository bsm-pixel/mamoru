-- 제품 발주명: 매입처에 주문 시 사용하는 이름 (TMS 제품명과 다를 수 있음)
ALTER TABLE products ADD COLUMN IF NOT EXISTS purchase_name text;
