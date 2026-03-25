-- 031: 계약서 필기 + 이미지 + 상담 연결
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS image_url text;
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS consultation_id uuid REFERENCES consultations(id);
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS handwriting_name text;
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS handwriting_phone text;
ALTER TABLE contracts ADD COLUMN IF NOT EXISTS handwriting_address text;
CREATE INDEX IF NOT EXISTS idx_contracts_consultation ON contracts(consultation_id);
