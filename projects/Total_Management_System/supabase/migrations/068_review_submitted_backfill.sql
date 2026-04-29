-- 068: 067 배포 이전에 작성된 리뷰들의 review_submitted_at 일괄 백필
-- 067 배포 후 작성되는 리뷰는 reviews/submit이 자동 매칭하므로 본 마이그는 1회용.
--
-- 매칭 기준 (source_id 정확 매칭, 오매칭 위험 0):
--   consultation: consultations.unique_id × reviews(type=consult).source_id
--   repair:       repairs.as_id          × reviews(type=repair).source_id
--   sale:         offline_sales.sale_number × reviews.source_id (type 무관)
--
-- 멱등: review_submitted_at IS NULL인 행만 업데이트. 여러 번 실행해도 안전.
-- Supabase SQL Editor에서 실행

-- 1) 상담관리 백필
UPDATE consultations c
SET    review_submitted_at = r.created_at
FROM   reviews r
WHERE  r.type = 'consult'
  AND  r.source_id = c.unique_id
  AND  c.review_submitted_at IS NULL;

-- 2) 복원수리 백필
UPDATE repairs rp
SET    review_submitted_at = r.created_at
FROM   reviews r
WHERE  r.type = 'repair'
  AND  r.source_id = rp.as_id
  AND  rp.review_submitted_at IS NULL;

-- 3) 오프라인 판매 백필 (sale_number 기준, consult/repair 타입 모두 매칭)
UPDATE offline_sales os
SET    review_submitted_at = r.created_at
FROM   reviews r
WHERE  r.source_id = os.sale_number
  AND  os.review_submitted_at IS NULL;

-- 결과 확인용 (실행 후 별도 SELECT로 확인 가능)
-- SELECT 'consultations' AS source, COUNT(*) FROM consultations WHERE review_submitted_at IS NOT NULL
-- UNION ALL
-- SELECT 'repairs', COUNT(*) FROM repairs WHERE review_submitted_at IS NOT NULL
-- UNION ALL
-- SELECT 'offline_sales', COUNT(*) FROM offline_sales WHERE review_submitted_at IS NOT NULL;
