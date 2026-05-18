-- ═══════════════════════════════════════════════════════════════
-- 마이그레이션 085 — 시리얼 product_id 보완 (sale_item_id 에서 역추출)
-- ───────────────────────────────────────────────────────────────
-- 배경: 082·083·084 마이그레이션 후 sale_item_id 는 100% 채워졌으나
--       product_id 가 NULL 인 시리얼 847건 발견.
-- 원인: 082 마이그레이션이 sale_item_id 만 채우고 product_id 안 채움.
--       083 은 sale_item_id IS NULL 조건으로 제외 → 보완 안 됨.
-- 작업: sale_item_id 가 채워진 시리얼의 product_id 를 그 sale_item 의
--       product_id 로 역추출하여 채움. 100% 정확 (외래키 정합성).
-- ═══════════════════════════════════════════════════════════════

-- ── STEP 1: product_id 일괄 보완 ──
UPDATE product_serials ps
SET product_id = osi.product_id
FROM offline_sale_items osi
WHERE ps.sale_item_id = osi.id
  AND ps.product_id IS NULL
  AND osi.product_id IS NOT NULL;

-- ── STEP 2: 최종 무결성 진단 (완전판) ──
-- sale_item_id 또는 product_id 둘 중 하나라도 NULL 인 시리얼
SELECT
  COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL) AS missing_sale_item_id,
  COUNT(*) FILTER (WHERE ps.product_id IS NULL) AS missing_product_id,
  COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL OR ps.product_id IS NULL) AS total_broken,
  COUNT(*) AS total_serials_with_sale,
  ROUND(100.0 * COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL OR ps.product_id IS NULL) / NULLIF(COUNT(*), 0), 2) AS broken_pct
FROM product_serials ps
WHERE ps.offline_sale_id IS NOT NULL;
