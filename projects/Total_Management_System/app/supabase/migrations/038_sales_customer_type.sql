-- 판매에 고객 유형 저장 (B2B/B2C 필터링용)
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS customer_type text DEFAULT NULL;

-- 기존 데이터 보정: customer_id가 있으면 customers.customer_type으로 채우기
UPDATE offline_sales s
SET customer_type = c.customer_type
FROM customers c
WHERE s.customer_id = c.id
  AND s.customer_type IS NULL;
