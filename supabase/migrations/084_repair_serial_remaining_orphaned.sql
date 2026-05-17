-- ═══════════════════════════════════════════════════════════════
-- 마이그레이션 084 — 잔여 orphaned 시리얼 강제 매핑 (각 판매 첫 sale_item)
-- ───────────────────────────────────────────────────────────────
-- 배경: 083 마이그레이션 후 818/870 → 2/870 (99.77% 복구).
--       잔여 2건은 시리얼 개수 > sale_items 개수 케이스 (잉여 시리얼).
--       정확한 매핑 정보 없음 → 각 판매의 첫 번째 sale_item에 강제 매핑.
-- 안전: NULL 인 시리얼만 처리 (이미 매핑된 시리얼은 영향 없음)
-- ═══════════════════════════════════════════════════════════════

-- ── STEP 1: 각 판매의 가장 작은 id sale_item을 강제 매핑 대상 ──
WITH first_items AS (
  SELECT DISTINCT ON (sale_id)
    id AS sale_item_id,
    sale_id,
    product_id
  FROM offline_sale_items
  ORDER BY sale_id, id
)
UPDATE product_serials ps
SET
  product_id = fi.product_id,
  sale_item_id = fi.sale_item_id
FROM first_items fi
WHERE ps.offline_sale_id = fi.sale_id
  AND ps.sale_item_id IS NULL;

-- ── STEP 2: 최종 결과 확인 (0건이어야 정상) ──
SELECT
  COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL) AS final_orphaned,
  COUNT(*) AS total,
  ROUND(100.0 * COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL) / NULLIF(COUNT(*), 0), 2) AS final_pct
FROM product_serials ps
WHERE ps.offline_sale_id IS NOT NULL;
