-- 정품확인 토큰 (QR에 노출되는 값, 시리얼번호 대신 사용)
ALTER TABLE product_serials ADD COLUMN IF NOT EXISTS verify_token text;
CREATE UNIQUE INDEX IF NOT EXISTS idx_serials_verify_token ON product_serials(verify_token) WHERE verify_token IS NOT NULL;
