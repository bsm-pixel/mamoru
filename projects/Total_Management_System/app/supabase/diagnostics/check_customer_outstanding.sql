-- ============================================================
-- SSOT 점검: customer.outstanding_balance ↔ 미결제 sale/delivery 합계
-- ============================================================
-- 목적: customers.outstanding_balance(저장값) 와
--       offline_sales + deliveries 의 미수금 합계(계산값) 가 일치하는지 점검.
--
-- 동작: READ-ONLY. 데이터 변경 없음.
-- 결과: diff != 0 인 고객만 노출. 0 rows 면 정합성 OK.
-- 사용: Supabase SQL Editor에서 그대로 RUN.
-- ============================================================

WITH sales_unpaid AS (
  SELECT customer_id,
         SUM(GREATEST(0, COALESCE(total_amount, 0)
                          - COALESCE(discount_amount, 0)
                          - COALESCE(paid_amount, 0))) AS amt
  FROM offline_sales
  WHERE customer_id IS NOT NULL
    AND payment_status IN ('unpaid', 'partial')
    AND cancelled_at IS NULL
    AND returned_at IS NULL
  GROUP BY customer_id
),
delivery_unpaid AS (
  SELECT customer_id,
         SUM(GREATEST(0, COALESCE(total_amount, 0)
                          - COALESCE(discount_amount, 0)
                          - COALESCE(paid_amount, 0))) AS amt
  FROM deliveries
  WHERE customer_id IS NOT NULL
    AND payment_status IN ('unpaid', 'partial')
    AND cancelled_at IS NULL
  GROUP BY customer_id
)
SELECT
  c.id,
  c.name,
  c.phone,
  c.customer_type,
  COALESCE(c.outstanding_balance, 0) AS stored,
  COALESCE(su.amt, 0) AS sales_unpaid,
  COALESCE(du.amt, 0) AS delivery_unpaid,
  COALESCE(su.amt, 0) + COALESCE(du.amt, 0) AS calculated,
  COALESCE(c.outstanding_balance, 0)
    - (COALESCE(su.amt, 0) + COALESCE(du.amt, 0)) AS diff
FROM customers c
LEFT JOIN sales_unpaid    su ON su.customer_id = c.id
LEFT JOIN delivery_unpaid du ON du.customer_id = c.id
WHERE COALESCE(c.outstanding_balance, 0)
   != COALESCE(su.amt, 0) + COALESCE(du.amt, 0)
ORDER BY ABS(
  COALESCE(c.outstanding_balance, 0)
  - (COALESCE(su.amt, 0) + COALESCE(du.amt, 0))
) DESC
LIMIT 100;

-- 해석:
--   diff > 0: 저장값이 실제 미수금보다 큼 (과대 표시) → 정합화 필요
--   diff < 0: 저장값이 실제 미수금보다 작음 (과소 표시) → 정합화 필요
--   diff = 0: 일치 (이 SQL은 0인 행은 출력 안 함)
--
-- 다음 단계 (정합화):
--   결과를 사장님이 검토 → 합의 후 별도 UPDATE SQL 작성하여
--   c.outstanding_balance := calculated 로 일괄 보정.
