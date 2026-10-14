-- 162: 판매 합포장 — 새 송장을 만들지 않고 다른 판매 건의 송장에 얹어 보내기 (2026-10-14)
--
-- 배경(사장님): 같은 매장(같은 주소)의 원장·직원이 각자 구매한 판매 건을 한 박스로 보내고 싶다.
--   납품(158)·복원수리엔 합포장이 있는데 판매끼리는 없었다. 「다른 택배사로 발송」은 롯데를 고를 수 없어
--   송장번호를 옮겨 적으면 자동추적·출고 알림톡이 전부 빠졌다.
--
-- 이 컬럼은 "이 판매는 저 판매의 송장에 얹혀 간다"는 표시다 (158 deliveries.merged_into_delivery_id 와 같은 규칙).
--   · 값이 있는 건 = 얹힌 건 (송장번호·택배사는 원 건과 같은 값이 복사돼 있다)
--   · 값이 없는 건 = 자기 송장을 가진 건(원 건) 또는 송장 없는 건
--
-- 출고완료·배송완료·알림톡은 기존 크론이 판매 건마다 송장번호로 조회하므로 그대로 함께 넘어간다.
--   → 각 판매 건의 고객 전화번호로 출고 알림톡·후기 요청이 따로 나간다.
-- 이 표시는 원 송장이 취소·수동 출고·판매 취소될 때 얹힌 건도 따라가게 하는 데 쓴다.
--   · 원 건 송장 취소 / 판매 취소 → 얹힌 건 송장도 함께 지움 (출고 전 건만)
--   · 원 건 수동 [출고완료]        → 얹힌 건도 함께 출고완료
--   · 얹힌 건 송장 취소            → 합포장만 풀림 (ALPS 취소 호출 안 함 — 원 건 송장이 살아 있다)

ALTER TABLE offline_sales
  ADD COLUMN IF NOT EXISTS merged_into_sale_id uuid REFERENCES offline_sales(id) ON DELETE SET NULL;

COMMENT ON COLUMN offline_sales.merged_into_sale_id IS
  '합포장: 이 판매가 얹혀 가는 원 송장 판매의 id. NULL=자기 송장(또는 송장 없음). 162';

CREATE INDEX IF NOT EXISTS offline_sales_merged_into_idx
  ON offline_sales (merged_into_sale_id)
  WHERE merged_into_sale_id IS NOT NULL;

-- 확인용
SELECT column_name, data_type FROM information_schema.columns
 WHERE table_name = 'offline_sales' AND column_name = 'merged_into_sale_id';
