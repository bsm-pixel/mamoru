-- 139_repairs_recall_pickup.sql — 복원수리 재수거(정밀 재점검) 롯데 반품수거 기록 (2026-08-27)
-- 출고된 복원수리 건을 고객집에서 다시 회수(재점검)할 때 롯데 반품접수(ustRtgSctCd=02) 송장을 기록.
-- 고객 노출 문구는 '정밀 재점검'(중립·긍정), 내부는 '재수거'. additive only.

ALTER TABLE repairs ADD COLUMN IF NOT EXISTS recall_invoice_number text;
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS recall_courier_name text;
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS recall_booked_at timestamptz;

COMMENT ON COLUMN repairs.recall_invoice_number IS '재수거(정밀 재점검) 롯데 송장(02, 취소는 ALPS 수동)';
