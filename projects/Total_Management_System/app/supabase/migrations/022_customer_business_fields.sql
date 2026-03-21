-- 022: 매입처/고객 사업자 정보 필드 추가
ALTER TABLE customers ADD COLUMN IF NOT EXISTS business_number text;      -- 사업자등록번호
ALTER TABLE customers ADD COLUMN IF NOT EXISTS representative text;       -- 대표자명
ALTER TABLE customers ADD COLUMN IF NOT EXISTS business_type text;        -- 업태
ALTER TABLE customers ADD COLUMN IF NOT EXISTS business_category text;    -- 종목
ALTER TABLE customers ADD COLUMN IF NOT EXISTS contact_channel text;      -- 연락 경로 (카톡, 전화 등)
