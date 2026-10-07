-- 158: 납품 합포장 — 새 송장을 만들지 않고 같은 거래처의 기존 송장에 얹어 보내기 (2026-10-07)
--
-- 배경: 송장을 만들어둔 납품(출고대기)이 있는데 출고 전에 같은 거래처 추가 주문이 들어온 경우,
--   한 박스에 같이 넣어 보낸다. 지금까진 「다른 택배사로 보냈어요」에 송장번호를 손으로 옮겨 적어야 했고
--   누르는 즉시 '출고완료'가 돼서, 기사님이 오기도 전에 한 건만 상태가 앞서갔다.
--
-- 이 컬럼은 "이 건은 저 납품의 송장에 얹혀 간다"는 표시다.
--   · 값이 있는 건 = 합포장으로 얹힌 건 (송장번호·택배사는 원 건과 같은 값이 복사돼 있다)
--   · 값이 없는 건 = 자기 송장을 가진 건(원 건) 또는 송장 없는 건
--
-- 상태 전환은 기존 크론이 건마다 송장번호로 조회하므로 그대로 함께 넘어간다(출고대기 → 출고완료 → 배송완료).
-- 이 표시는 "원 송장이 취소·재발급·수동 출고될 때 얹힌 건도 따라가게" 하는 데 쓴다.
--   · 원 건 송장 취소 / 납품 취소 → 얹힌 건 송장도 함께 지움
--   · 원 건 송장 재발급            → 얹힌 건도 새 번호로
--   · 원 건 수동 [출고 완료]        → 얹힌 건도 함께 출고완료
--   · 얹힌 건 송장 취소             → 합포장만 풀림(원 건 무영향)

ALTER TABLE deliveries
  ADD COLUMN IF NOT EXISTS merged_into_delivery_id uuid REFERENCES deliveries(id) ON DELETE SET NULL;

COMMENT ON COLUMN deliveries.merged_into_delivery_id IS
  '합포장: 이 납품이 얹혀 가는 원 송장 납품의 id. NULL=자기 송장(또는 송장 없음)';

CREATE INDEX IF NOT EXISTS deliveries_merged_into_idx
  ON deliveries (merged_into_delivery_id)
  WHERE merged_into_delivery_id IS NOT NULL;

-- 확인용
SELECT column_name, data_type FROM information_schema.columns
 WHERE table_name = 'deliveries' AND column_name = 'merged_into_delivery_id';
