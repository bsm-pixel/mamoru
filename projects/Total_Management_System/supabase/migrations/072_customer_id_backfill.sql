-- 072: customer_id 백필 (옛 NULL 데이터 → phone 매칭으로 자동 연결)
-- 신규 데이터는 lib/customer/match-or-create.ts로 자동 채워지므로 1회성 정리
-- Supabase SQL Editor에서 실행
--
-- 안전성:
--   - customer_id NULL인 행만 대상 (멱등)
--   - 같은 phone 여러 customers 있으면 가장 오래된(첫 등록) customer로 매칭 (deterministic)
--   - phone_normalized 비어있으면 skip

-- 1) consultations
UPDATE consultations c
SET    customer_id = (
  SELECT id FROM customers
  WHERE  phone_normalized = c.phone_normalized
  ORDER BY created_at ASC
  LIMIT 1
)
WHERE  c.customer_id IS NULL
  AND  c.phone_normalized IS NOT NULL
  AND  c.phone_normalized != ''
  AND  EXISTS (SELECT 1 FROM customers WHERE phone_normalized = c.phone_normalized);

-- 2) repairs
UPDATE repairs r
SET    customer_id = (
  SELECT id FROM customers
  WHERE  phone_normalized = r.phone_normalized
  ORDER BY created_at ASC
  LIMIT 1
)
WHERE  r.customer_id IS NULL
  AND  r.phone_normalized IS NOT NULL
  AND  r.phone_normalized != ''
  AND  EXISTS (SELECT 1 FROM customers WHERE phone_normalized = r.phone_normalized);

-- 3) offline_sales — customer_phone에서 정규화 후 매칭 (offline_sales는 phone_normalized 컬럼 없음)
UPDATE offline_sales s
SET    customer_id = (
  SELECT id FROM customers
  WHERE  phone_normalized = REGEXP_REPLACE(COALESCE(s.customer_phone, ''), '\D', '', 'g')
  ORDER BY created_at ASC
  LIMIT 1
)
WHERE  s.customer_id IS NULL
  AND  s.customer_phone IS NOT NULL
  AND  EXISTS (
    SELECT 1 FROM customers
    WHERE  phone_normalized = REGEXP_REPLACE(COALESCE(s.customer_phone, ''), '\D', '', 'g')
  );

-- 결과 확인용 (실행 후 별도 SELECT로)
-- SELECT 'consultations' AS t, COUNT(*) FILTER (WHERE customer_id IS NOT NULL) AS linked, COUNT(*) AS total FROM consultations
-- UNION ALL
-- SELECT 'repairs', COUNT(*) FILTER (WHERE customer_id IS NOT NULL), COUNT(*) FROM repairs
-- UNION ALL
-- SELECT 'offline_sales', COUNT(*) FILTER (WHERE customer_id IS NOT NULL), COUNT(*) FROM offline_sales;
