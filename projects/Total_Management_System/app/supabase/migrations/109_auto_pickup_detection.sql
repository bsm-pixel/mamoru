-- 109: 롯데 집하(수거) 자동 감지 → 자동 출고완료 + B2C 출고 알림톡
--
-- 배경 (2026-07-12):
--   롯데 기사님이 방문 수거하며 스캔하면 ALPS 화물상태가 '집하'(godsStatCd '10')로 바뀐다.
--   그 코드는 이미 추적 API 응답에 들어오고 있었으나 TMS 가 09/41/45 만 보고 나머지를 버리고 있었다.
--   → 집하 감지 시점에 자동으로 출고완료 처리 + (B2C 고객에게만) 출고 알림톡 발송.
--   사장님의 수동 클릭 2회([출고완료] + 알림톡 체크)가 0회가 된다.
--
-- 변경:
--   offline_sales.shipped_source       'manual' | 'alps_pickup'  (NULL = 레거시/수동)
--   offline_sales.shipped_notified_at  출고 알림톡(sales_shipped) 발송 성공 시각. B2C 만.
--   repairs.shipped_source             동일 (복원수리도 집하 자동 감지 대상)
--
-- 영향:
--   기존 데이터 무영향 (전부 nullable, 백필 없음 → 기존 행은 NULL = 수동으로 취급).
--   기존 쿼리 무영향.

ALTER TABLE offline_sales
  ADD COLUMN IF NOT EXISTS shipped_source      TEXT,
  ADD COLUMN IF NOT EXISTS shipped_notified_at TIMESTAMPTZ;

COMMENT ON COLUMN offline_sales.shipped_source      IS '출고완료 기록 주체. manual=사장님 버튼, alps_pickup=롯데 집하 스캔 자동감지(cron)';
COMMENT ON COLUMN offline_sales.shipped_notified_at IS '출고 알림톡(sales_shipped) 발송 성공 시각. B2C 만 발송. NULL=미발송';

ALTER TABLE repairs
  ADD COLUMN IF NOT EXISTS shipped_source TEXT;

COMMENT ON COLUMN repairs.shipped_source IS '출고 기록 주체. manual=사장님 상태변경, alps_pickup=롯데 집하 스캔 자동감지(cron)';

-- 집하 감지 폴링 대상 최적화 (partial index)
--   조건: 송장 있음 + 아직 출고 전 + 배송 미완료 + 미취소
--   ORDER BY sale_date DESC 로 최신건 우선 폴링 (적체 방지)
CREATE INDEX IF NOT EXISTS idx_offline_sales_awaiting_pickup
  ON offline_sales(sale_date DESC)
  WHERE shipped_at IS NULL
    AND delivered_at IS NULL
    AND cancelled_at IS NULL
    AND invoice_number IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_repairs_awaiting_pickup
  ON repairs(invoice_number)
  WHERE status = 'ready_to_ship'
    AND invoice_number IS NOT NULL;

-- 검증 SQL (사장님이 Supabase SQL Editor 에서 직접 확인)
-- SELECT column_name, data_type, is_nullable FROM information_schema.columns
-- WHERE (table_name='offline_sales' AND column_name IN ('shipped_source','shipped_notified_at'))
--    OR (table_name='repairs'       AND column_name = 'shipped_source');
-- → 3행 (전부 YES = nullable)
--
-- 자동 감지된 건 확인 (배포 후)
-- SELECT sale_number, customer_name, customer_type, shipped_at, shipped_source, shipped_notified_at
-- FROM offline_sales WHERE shipped_source='alps_pickup' ORDER BY shipped_at DESC LIMIT 20;
--
-- B2B 에 잘못 나간 알림톡이 없는지 (0행이어야 정상)
-- SELECT sale_number FROM offline_sales
-- WHERE shipped_notified_at IS NOT NULL AND customer_type IN ('dealer','academy');
