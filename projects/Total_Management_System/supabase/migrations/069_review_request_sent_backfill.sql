-- 069: 067 배포 이전에 자동 발송된 후기 요청의 review_request_sent_at 일회성 백필
-- 067 배포 이전: 상담완료('completed')/수리배송완료('delivered') 전환 시 자동 알림톡 발송 동작
-- 그러나 review_request_sent_at 컬럼이 067에서 신설되어 발송 시점 미기록 → '미발송' 칩 오표시
-- → 사장님이 중복 발송 위험. history 기반으로 백필.
--
-- 매칭 기준:
--   consultation: consultation_history.to_status='completed'의 MIN(created_at)
--   repair:       repair_history.to_status='delivered'의 MIN(created_at)
--
-- 멱등: review_request_sent_at IS NULL인 행만 UPDATE.
-- 부정확 가능성 (낮음): phone 없거나 발송 실패한 케이스도 백필됨 — 사장님 인지 필요.
--   (실제 발송 여부는 Make 시나리오 / 솔라피 발송 이력에서 확인 가능)
-- Supabase SQL Editor에서 실행

-- 1) 상담관리: completed 전환 시점으로 백필 (review_request 템플릿은 모든 consultation_type에 발송됨)
UPDATE consultations c
SET    review_request_sent_at = h.first_completed_at
FROM (
  SELECT consultation_id, MIN(created_at) AS first_completed_at
  FROM   consultation_history
  WHERE  to_status = 'completed'
  GROUP BY consultation_id
) h
WHERE  c.id = h.consultation_id
  AND  c.status = 'completed'                  -- 현재도 completed인 건만 (취소·롤백 케이스 제외)
  AND  c.review_request_sent_at IS NULL;

-- 2) 복원수리: delivered 전환 시점으로 백필
UPDATE repairs rp
SET    review_request_sent_at = h.first_delivered_at
FROM (
  SELECT repair_id, MIN(created_at) AS first_delivered_at
  FROM   repair_history
  WHERE  to_status = 'delivered'
  GROUP BY repair_id
) h
WHERE  rp.id = h.repair_id
  AND  rp.status = 'delivered'                 -- 현재도 delivered인 건만
  AND  rp.review_request_sent_at IS NULL;

-- offline_sales: 067 이전에 자동 발송 흐름 없었음 (이미 ReviewRequestModal로 수동) →
--                review_requested_at(legacy alias)이 정확하므로 백필 대상 아님

-- 결과 확인용
-- SELECT 'consultations' AS source, COUNT(*) AS sent_count
-- FROM   consultations WHERE review_request_sent_at IS NOT NULL
-- UNION ALL
-- SELECT 'repairs', COUNT(*) FROM repairs WHERE review_request_sent_at IS NOT NULL;
