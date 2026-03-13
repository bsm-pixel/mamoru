-- 011: 제품구매 리뷰 파이프라인 — orders에 review_requested_at 추가
ALTER TABLE orders ADD COLUMN IF NOT EXISTS review_requested_at timestamptz;
