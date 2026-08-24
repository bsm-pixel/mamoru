-- 130: 고객 활동유형 집계 RPC (phone 기준) — 목록 칩용 (2026-08-23)
--
-- 정규화 전화번호 배열을 받아, 각 번호가 아임웹주문/오프라인구매/복원수리 활동이 있는지 반환.
-- 목록 화면(주문·판매·복원수리·고객·상담)에서 한 고객 옆에 활동유형 칩을 붙이는 데이터소스.
-- customer_id 가 아직 안 채워진 도메인(orders)도 있어, 가장 견고한 phone_normalized 기준으로 집계.
-- (repairs/consultations 는 phone_normalized 컬럼 사용, orders/offline_sales 는 즉시 정규화 비교)

CREATE OR REPLACE FUNCTION get_activity_types_by_phones(phones text[])
RETURNS TABLE(phone text, has_order boolean, has_sale boolean, has_repair boolean)
LANGUAGE sql STABLE AS $$
  SELECT p AS phone,
    EXISTS(
      SELECT 1 FROM orders o
      WHERE regexp_replace(coalesce(o.orderer_phone, ''), '\D', '', 'g') = p
        AND o.status NOT IN ('cancelled', 'refunded')
    ) AS has_order,
    EXISTS(
      SELECT 1 FROM offline_sales s
      WHERE regexp_replace(coalesce(s.customer_phone, ''), '\D', '', 'g') = p
        AND s.cancelled_at IS NULL
    ) AS has_sale,
    EXISTS(
      SELECT 1 FROM repairs r
      WHERE r.phone_normalized = p
    ) AS has_repair
  FROM unnest(phones) AS p
  WHERE p IS NOT NULL AND p <> '';
$$;
