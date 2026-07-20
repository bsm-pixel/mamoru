-- 111: 판매 '포장완료(준비완료)' 표시 — 2026-07-18
--
-- 배경: 판매관리 상태가 '준비중'(송장 발급됨, 출고 전) 하나로 뭉쳐 있어
--       "물건 챙겨서 포장까지 끝냈는지"를 표시할 자리가 없었다.
--       사장님 요청: 포장 끝낸 건은 '준비완료'로 구분해서 보이게.
--
-- 설계: 복원수리(repairs.packed_at)와 **동일한 개념·동일한 컬럼명**으로 통일.
--       · 송장 유무와 무관 (포장은 송장 발급 전에도 함) — 사장님 확정
--       · 내부 표시 전용 — 알림톡/외부연동 없음
--       · 상태 파생 우선순위: delivered > shipped > packed > invoice > 미처리
--
-- 되돌리기: ALTER TABLE offline_sales DROP COLUMN packed_at;

ALTER TABLE offline_sales
  ADD COLUMN IF NOT EXISTS packed_at timestamptz;

COMMENT ON COLUMN offline_sales.packed_at IS
  '포장완료(준비완료) 시점. NULL=미포장. 송장 유무와 무관한 내부 준비 상태 표시 (111, 2026-07-18)';

-- 준비완료 건만 빠르게 거르기 위한 부분 인덱스
-- (출고 전 + 포장 끝난 건 = 사장님이 "기사님만 오면 되는" 목록)
CREATE INDEX IF NOT EXISTS idx_offline_sales_packed_pending
  ON offline_sales (packed_at)
  WHERE packed_at IS NOT NULL AND shipped_at IS NULL;
