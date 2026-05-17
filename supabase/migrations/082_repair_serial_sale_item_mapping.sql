-- ═══════════════════════════════════════════════════════════════
-- 마이그레이션 082 — 기존 깨진 시리얼 sale_item_id 일괄 복구
-- ───────────────────────────────────────────────────────────────
-- 배경: 2026-05-17 진단 결과 product_serials.sale_item_id 가 93.92%
--       (818/871건) NULL 상태. rebuild_sale API 결함 + 레거시 데이터 누적.
-- 작업: 단순 케이스(한 판매에 한 상품 = 한 줄)는 자동 복구.
--       복잡 케이스(같은 상품 여러 줄)는 별도 처리 필요 — 본 마이그레이션 후
--       남는 orphaned 시리얼은 사장님 수동 검증.
-- 코드 fix 동반: api/sales/[id]/route.ts STEP 3.5 재매칭 로직 재작성 (커밋 동시)
-- ═══════════════════════════════════════════════════════════════

-- ── STEP 1: 단순 케이스 자동 복구 ──────────────────────────────
-- 조건: 해당 판매에 그 상품(product_id) 줄이 단 1개일 때만 매칭
--       (다중 줄이면 어느 줄에 매핑할지 모호 → 보류)

WITH single_match AS (
  SELECT
    ps.id AS serial_id,
    osi.id AS sale_item_id
  FROM product_serials ps
  JOIN offline_sale_items osi
    ON osi.sale_id = ps.offline_sale_id
    AND osi.product_id = ps.product_id
  WHERE ps.offline_sale_id IS NOT NULL
    AND ps.sale_item_id IS NULL
    AND (
      SELECT COUNT(*)
      FROM offline_sale_items osi2
      WHERE osi2.sale_id = ps.offline_sale_id
        AND osi2.product_id = ps.product_id
    ) = 1
)
UPDATE product_serials ps
SET sale_item_id = sm.sale_item_id
FROM single_match sm
WHERE ps.id = sm.serial_id;

-- ── STEP 2: 결과 확인 (사장님이 실행 후 결과 보내주세요) ──────
-- 복구 후 남은 orphaned 카운트
SELECT
  COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL) AS orphaned_after,
  COUNT(*) AS total,
  ROUND(100.0 * COUNT(*) FILTER (WHERE ps.sale_item_id IS NULL) / NULLIF(COUNT(*), 0), 2) AS orphaned_pct_after
FROM product_serials ps
WHERE ps.offline_sale_id IS NOT NULL;

-- ── STEP 3: 복잡 케이스 (수동 검증) ────────────────────────────
-- 같은 판매에 같은 product_id 가 여러 줄로 있는 경우 — 어느 줄에 매핑할지 모호
-- 아래 SELECT 로 목록 확인 후 사장님이 직접 검증해서 UPDATE
SELECT
  os.sale_number,
  os.created_at::date,
  c.name AS customer,
  ps.product_id,
  COUNT(ps.id) AS orphaned_serials,
  COUNT(DISTINCT osi.id) AS sale_item_lines
FROM product_serials ps
JOIN offline_sales os ON os.id = ps.offline_sale_id
LEFT JOIN customers c ON c.id = os.customer_id
JOIN offline_sale_items osi
  ON osi.sale_id = ps.offline_sale_id
  AND osi.product_id = ps.product_id
WHERE ps.offline_sale_id IS NOT NULL
  AND ps.sale_item_id IS NULL
  AND os.cancelled_at IS NULL
GROUP BY os.id, os.sale_number, os.created_at, c.name, ps.product_id
HAVING COUNT(DISTINCT osi.id) > 1
ORDER BY os.created_at DESC
LIMIT 50;
