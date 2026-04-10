-- 060: 고객 태그 컬럼 추가
-- Supabase SQL Editor에서 실행

ALTER TABLE customers
  ADD COLUMN IF NOT EXISTS tags text[] DEFAULT '{}';

-- GIN 인덱스: 태그 기반 필터링 성능 최적화
CREATE INDEX IF NOT EXISTS idx_customers_tags ON customers USING GIN (tags);
