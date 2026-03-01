-- 010: 고객 검색 인덱스 (자동완성용 trigram 부분매칭)
-- Supabase SQL Editor에서 실행

-- pg_trgm: ILIKE 부분매칭을 위한 GIN 인덱스 지원
CREATE EXTENSION IF NOT EXISTS pg_trgm;

-- 고객명 trigram 인덱스 (ILIKE %검색어% 최적화)
CREATE INDEX IF NOT EXISTS idx_customers_name_trgm
  ON customers USING gin (name gin_trgm_ops);

-- 전화번호(정규화) trigram 인덱스
CREATE INDEX IF NOT EXISTS idx_customers_phone_trgm
  ON customers USING gin (phone_normalized gin_trgm_ops);
