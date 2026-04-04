-- 052: 제품 B2B 납품명 (딜러/아카데미별 별도 제품명)
-- Supabase SQL Editor에서 실행

ALTER TABLE products ADD COLUMN IF NOT EXISTS dealer_name text;
ALTER TABLE products ADD COLUMN IF NOT EXISTS academy_name text;
