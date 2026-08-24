-- 131: 기존 아임웹 주문(orders)을 TMS 고객(customers)과 연결 — customer_id backfill (2026-08-23)
--
-- 그동안 sync 가 orders.customer_id 를 안 채워 항상 NULL 이었다(고객 상세뷰에 주문이 안 붙던 원인).
-- 앞으로는 sync 가 matchOrCreate 로 채우고, 과거 주문은 이 backfill 로 기존 고객에 연결한다.
-- 매칭 기준: 정규화 전화번호 == 고객 phone_normalized (병합 숨김 고객 제외).
-- ※ 해당 전화번호의 고객이 아직 없으면 NULL 유지(온라인 전용 고객) — 향후 다른 접점에서 생성 시 자동 연결.

UPDATE orders o
SET customer_id = c.id
FROM customers c
WHERE o.customer_id IS NULL
  AND c.merged_into_id IS NULL
  AND regexp_replace(coalesce(o.orderer_phone, ''), '\D', '', 'g') <> ''
  AND c.phone_normalized = regexp_replace(coalesce(o.orderer_phone, ''), '\D', '', 'g');
