-- 120: 복원수리 결제수단 (직접방문 현장결제 이체/카드/현금 + 회계 결제수단별 집계 반영)
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS payment_method text;
COMMENT ON COLUMN repairs.payment_method IS '결제수단 transfer/card/cash (직접방문 현장결제·입금확인 시 기록, 리포트 by_method A채널 반영, 120 2026-07-28)';
