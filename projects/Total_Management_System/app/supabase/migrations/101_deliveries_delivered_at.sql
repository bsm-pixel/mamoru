-- ╔════════════════════════════════════════════════════════════════════╗
-- ║ 101_deliveries_delivered_at — B2B 납품 배송완료 자동추적             ║
-- ║                                                                    ║
-- ║ 배경 (2026-06-01 사장님): 고객(B2C offline_sales)은 인수자등록 시    ║
-- ║   배송완료일이 뜨는데, 거래처(B2B deliveries)는 추적에서 빠져있음.   ║
-- ║   → deliveries 도 ALPS 인수자등록(코드 45) 감지 시 배송완료 표시.    ║
-- ║   ※ deliveries 엔 이미 tracking_number(송장) 있음 → 그걸로 ALPS 조회.║
-- ║                                                                    ║
-- ║ ⚠️ ADDITIVE ONLY — nullable. 기존 납품 데이터/흐름 영향 0.           ║
-- ╚════════════════════════════════════════════════════════════════════╝

ALTER TABLE deliveries ADD COLUMN IF NOT EXISTS delivered_at TIMESTAMPTZ; -- 배송완료(인수자등록) 시각

-- 자동추적 대상(송장 있고 미배송) 빠른 조회용
CREATE INDEX IF NOT EXISTS idx_deliveries_inflight
  ON deliveries(tracking_number)
  WHERE tracking_number IS NOT NULL AND delivered_at IS NULL;
