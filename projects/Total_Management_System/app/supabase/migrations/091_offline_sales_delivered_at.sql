-- 091: offline_sales 자동 배송완료 추적 + 고객 수령 완료 처리
--
-- 배경 (2026-05-25):
--   사장님 비전 — 판매관리도 자동 배송완료 추적 + 후기요청 자동 발송 + 매장 직접 수령 처리.
--   3채널 통합 IA (repairs / orders / offline_sales) — 송장 있는 모든 흐름 자동 추적.
--
-- 변경:
--   offline_sales 에 delivered_at timestamptz 컬럼 추가 (NULL 허용).
--   - 자동 cron (ALPS 인수자등록 감지): delivered_at = 감지 시각
--   - 수동 "고객 수령 완료" 버튼 (송장 없는 매장수령): delivered_at = 클릭 시각
--   - invoice_number 유무로 배송완료 vs 수령완료 자동 분기 (UI 칩은 '판매완료' 단일 통일)
--
-- 영향:
--   기존 데이터 무영향 (nullable). 기존 쿼리 무영향.

ALTER TABLE offline_sales
  ADD COLUMN IF NOT EXISTS delivered_at TIMESTAMPTZ;

COMMENT ON COLUMN offline_sales.delivered_at IS '배송완료/고객수령 시각. 자동(cron ALPS 인수자등록) 또는 수동(매장 직접 수령 버튼)';

-- 자동 추적 대상 쿼리 최적화 (partial index)
--   조건: 송장 있음 + 발송 완료 + 배송 미완료 + 미취소
CREATE INDEX IF NOT EXISTS idx_offline_sales_inflight
  ON offline_sales(shipped_at)
  WHERE delivered_at IS NULL
    AND cancelled_at IS NULL
    AND invoice_number IS NOT NULL;

-- 검증 SQL (사장님이 Supabase SQL Editor 에서 직접 확인)
-- SELECT column_name, data_type, is_nullable
-- FROM information_schema.columns
-- WHERE table_name='offline_sales' AND column_name='delivered_at';
-- → 1행: delivered_at, timestamp with time zone, YES
