-- 복원수리: 입금확인을 파이프라인에서 분리 → paid_at 독립 플래그
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS paid_at TIMESTAMPTZ DEFAULT NULL;

-- 기존 payment_confirmed 이후 상태 데이터 마이그레이션
UPDATE repairs SET paid_at = updated_at
WHERE status IN ('payment_confirmed','repairing','ready_to_ship','shipped','delivered','completed')
  AND paid_at IS NULL;

-- 현재 payment_confirmed 상태 → repairing으로 일괄 전환
UPDATE repairs SET status = 'repairing' WHERE status = 'payment_confirmed';
