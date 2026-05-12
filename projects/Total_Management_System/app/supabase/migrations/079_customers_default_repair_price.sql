-- 079: customers.default_repair_price — 거래처별 복원수리 기본 단가 (자루당, 원)
--
-- 배경 (2단계 2-E, 2026-05-12):
--   납품("+B2B수리") 입력 시 거래처마다 복원수리 단가가 다른데 매번 손으로 입력해야 했음.
--   거래처 마스터에 기본 단가를 저장해 두면 거래처 선택 시 자동으로 채워준다.
--   NULL 이면 입력 화면 기본값(8,000원) fallback. (B2B 거래처 = customer_type IN ('dealer','academy') 에 주로 사용)
ALTER TABLE customers ADD COLUMN IF NOT EXISTS default_repair_price integer;
COMMENT ON COLUMN customers.default_repair_price IS '거래처별 복원수리 기본 단가(자루당, 원). NULL = 입력 화면 기본값(8000) fallback';
