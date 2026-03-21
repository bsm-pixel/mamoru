-- Phase 0: 아카데미 전용 가격 컬럼 추가
ALTER TABLE products ADD COLUMN IF NOT EXISTS price_academy bigint DEFAULT 0;
