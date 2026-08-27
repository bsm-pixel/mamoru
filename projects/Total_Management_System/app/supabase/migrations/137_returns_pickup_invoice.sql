-- 137_returns_pickup_invoice.sql — 반품 수거접수(롯데 반품 API) 송장 기록 (2026-08-27)
-- 롯데 IS팀 회신: ustRtgSctCd='02'=반품, orglInvNo=원송장(선택). 취소 API는 미지원.
-- 반품 수거접수 시 발급된 롯데 송장번호·집하일을 returns 에 기록한다. additive only.

ALTER TABLE returns ADD COLUMN IF NOT EXISTS pickup_invoice_number text;
ALTER TABLE returns ADD COLUMN IF NOT EXISTS pickup_courier_name text;
ALTER TABLE returns ADD COLUMN IF NOT EXISTS pickup_booked_at timestamptz;

COMMENT ON COLUMN returns.pickup_invoice_number IS '반품 수거접수 롯데 송장번호(ustRtgSctCd=02, 취소는 ALPS 수동)';
COMMENT ON COLUMN returns.pickup_booked_at IS '반품 수거접수 시각';
