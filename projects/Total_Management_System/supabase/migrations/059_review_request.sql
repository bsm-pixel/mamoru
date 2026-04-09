-- 059: 판매 건에서 후기 요청 발송 추적
ALTER TABLE offline_sales
  ADD COLUMN IF NOT EXISTS review_requested_at timestamptz;
