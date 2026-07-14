-- 110: B2B 납품 — "송장 생성 = 출고완료" 오류 수정 + 집하 자동 감지
--
-- 배경 (2026-07-12, 사장님 지적):
--   납품은 송장만 만들었는데 화면에 "출고완료"가 떴다. 기사님은 아직 오지도 않았는데.
--   원인: api/lotte/book 이 송장 발급과 동시에 deliveries.status='shipped' 를 강제.
--   B2C 판매는 송장번호만 넣고 출고는 집하 감지가 채우는데(109), B2B만 다르게 구현돼 있었다.
--
-- 변경 (109 판매와 동일한 정의로 통일):
--   송장 발급 = '출고대기' (status='confirmed' 유지 + tracking_number 만 저장)
--   기사님 수거(집하) 감지 → status='shipped' + shipped_date + shipped_source='alps_pickup'
--   deliveries.shipped_source  'manual' | 'alps_pickup'  (NULL = 레거시/수동)
--
-- 영향:
--   🟢 매출 집계: hub_stats RPC(077/078/080/088)·reports/summary 가 status IN ('confirmed','shipped','settled')
--      → confirmed 가 이미 포함되므로 매출 변동 0
--   🟢 미수금: lib/outstanding.ts 는 payment_status 만 봄 (status 무관)
--   🟢 결제완료 처리: update_payment 액션은 status 가드 없음 → 출고 전에도 결제 처리 가능
--   🟢 기존 데이터: nullable 컬럼 추가만, 백필 없음

ALTER TABLE deliveries
  ADD COLUMN IF NOT EXISTS shipped_source TEXT;

COMMENT ON COLUMN deliveries.shipped_source IS '출고 기록 주체. manual=사장님 [출고완료] 버튼, alps_pickup=롯데 집하 스캔 자동감지(cron)';

-- 집하 감지 폴링 대상 (송장 발급됐지만 아직 수거 전)
CREATE INDEX IF NOT EXISTS idx_deliveries_awaiting_pickup
  ON deliveries(tracking_number)
  WHERE status = 'confirmed'
    AND tracking_number IS NOT NULL
    AND cancelled_at IS NULL;


-- ══════════════════════════════════════════════════════════════════════
-- 기존 데이터 정정 (⚠️ 사장님이 미리보기 확인 후 실행)
-- ══════════════════════════════════════════════════════════════════════
--
-- "송장 생성으로 잘못 shipped 된 건" 을 정확히 식별할 수 있다:
--   · api/lotte/book 경로는 shipped_date 를 안 채운다  → shipped_date IS NULL
--   · 수동 [출고 완료] 버튼은 shipped_date 를 채운다     → shipped_date NOT NULL
--
-- ① 되돌릴 대상 미리보기 (실행 전 반드시 눈으로 확인)
-- SELECT dl_number, customer_name, tracking_number, status, shipped_date, delivered_at, created_at
-- FROM deliveries
-- WHERE status = 'shipped' AND shipped_date IS NULL AND delivered_at IS NULL AND cancelled_at IS NULL
-- ORDER BY created_at DESC;
--
-- ② 이 중 **아직 기사님이 안 가져간 건만** 출고대기로 되돌린다.
--    ⚠️ 실제로 집하된 건(예: DL-20260712-001)은 제외할 것 — dl_number 로 직접 지정 권장.
-- UPDATE deliveries SET status = 'confirmed', updated_at = now()
-- WHERE dl_number IN ('DL-20260714-001');   -- ← 되돌릴 건만 나열
--
-- ③ 매출 불변 확인 (②의 전후로 합계가 같아야 정상 — RPC 가 confirmed 포함하므로)
-- SELECT status, count(*), sum(total_amount - COALESCE(discount_amount, 0))
-- FROM deliveries WHERE cancelled_at IS NULL GROUP BY status ORDER BY status;
--
-- ④ 검증: 컬럼 추가 확인
-- SELECT column_name, data_type, is_nullable FROM information_schema.columns
-- WHERE table_name = 'deliveries' AND column_name = 'shipped_source';
