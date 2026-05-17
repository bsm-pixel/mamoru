-- ═══════════════════════════════════════════════════════════════
-- 마이그레이션 083 — 시리얼 product_id + sale_item_id 동시 복구
-- ───────────────────────────────────────────────────────────────
-- 배경: 082 진단 결과 깨진 시리얼 818건 *전부* product_id 도 NULL.
--       즉 시리얼이 "어느 상품의 시리얼인지" 정보 자체 없음.
--       offline_sale_id → offline_sale_items.product_id 로 역추출 가능.
-- 작업: 한 판매에 sale_item 줄이 1개일 때 자동 복구 (100% 정확)
--       여러 줄일 때는 다음 단계 (사장님 수동 검증 또는 ROW_NUMBER 순서 분배)
-- ═══════════════════════════════════════════════════════════════

-- ── STEP 1: 단순 케이스 자동 복구 (한 판매 = 한 sale_item 줄) ──
-- product_id + sale_item_id 동시 채움

WITH single_line AS (
  SELECT
    ps.id AS serial_id,
    osi.id AS sale_item_id,
    osi.product_id
  FROM product_serials ps
  JOIN offline_sale_items osi ON osi.sale_id = ps.offline_sale_id
  WHERE ps.offline_sale_id IS NOT NULL
    AND ps.sale_item_id IS NULL
    AND (
      SELECT COUNT(*) FROM offline_sale_items osi2
      WHERE osi2.sale_id = ps.offline_sale_id
    ) = 1
)
UPDATE product_serials ps
SET
  product_id = sl.product_id,
  sale_item_id = sl.sale_item_id
FROM single_line sl
WHERE ps.id = sl.serial_id;

-- ── STEP 2: 결과 확인 ──
SELECT
  COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL) AS still_orphaned,
  COUNT(*) AS total,
  ROUND(100.0 * COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL) / NULLIF(COUNT(*), 0), 2) AS pct_after
FROM product_serials ps
WHERE ps.offline_sale_id IS NOT NULL;

-- ── STEP 3: 다중 줄 케이스 — 순서대로 1:1 분배 (보수적) ──
-- 한 판매에 sale_item 여러 줄 + 시리얼 여러 개일 때
-- 시리얼을 created_at 순서로 + sale_items를 id 순서로 분배 (각 줄 1시리얼 가정)
-- ⚠️ 정확도 100% 아님 — 사장님이 잘못된 매칭 발견 시 수동 수정

WITH
  serials_ordered AS (
    SELECT
      ps.id AS serial_id,
      ps.offline_sale_id,
      ROW_NUMBER() OVER (PARTITION BY ps.offline_sale_id ORDER BY ps.created_at, ps.id) AS rn
    FROM product_serials ps
    WHERE ps.offline_sale_id IS NOT NULL
      AND ps.sale_item_id IS NULL
  ),
  items_ordered AS (
    SELECT
      osi.id AS sale_item_id,
      osi.sale_id,
      osi.product_id,
      ROW_NUMBER() OVER (PARTITION BY osi.sale_id ORDER BY osi.id) AS rn
    FROM offline_sale_items osi
    WHERE osi.sale_id IN (SELECT DISTINCT offline_sale_id FROM serials_ordered)
  )
UPDATE product_serials ps
SET
  product_id = io.product_id,
  sale_item_id = io.sale_item_id
FROM serials_ordered so
JOIN items_ordered io
  ON io.sale_id = so.offline_sale_id
  AND io.rn = so.rn
WHERE ps.id = so.serial_id
  AND ps.sale_item_id IS NULL;

-- ── STEP 4: 최종 결과 확인 ──
SELECT
  COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL) AS final_orphaned,
  COUNT(*) AS total,
  ROUND(100.0 * COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL) / NULLIF(COUNT(*), 0), 2) AS final_pct
FROM product_serials ps
WHERE ps.offline_sale_id IS NOT NULL;
