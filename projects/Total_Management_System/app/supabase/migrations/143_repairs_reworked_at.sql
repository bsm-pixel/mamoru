-- 143: 복원수리 재작업 시작 시각 (재수리 탭 분기용)
-- 재수거 접수(recall_booked_at)만으론 '재작업 대기'와 '재출고 완료'를 status로 구분 불가
-- (둘 다 shipped 가능) → 재작업 시작 시점을 별도 마커로 기록.
--   재수리 탭        = recall_booked_at IS NOT NULL AND reworked_at IS NULL   (재수거했으나 재작업 전)
--   출고완료 탭       = ... AND (recall_booked_at IS NULL OR reworked_at IS NOT NULL)  (재출고 완료건은 복귀)
ALTER TABLE repairs
  ADD COLUMN IF NOT EXISTS reworked_at TIMESTAMPTZ;

-- 재수리 탭 카운트 조회 최적화 (재수거 있고 재작업 전인 건만)
CREATE INDEX IF NOT EXISTS idx_repairs_recall_pending
  ON repairs(recall_booked_at)
  WHERE recall_booked_at IS NOT NULL AND reworked_at IS NULL;
