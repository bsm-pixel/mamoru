-- 159: 별도 배송지 — 고객/거래처 등록 주소가 아닌 곳으로 보낸 건 기록 (2026-10-10)
--
-- 배경(사장님): 거래처 이병관처럼 집 주소가 등록돼 있는데, 받는 곳을 그때그때 다르게 적어온다.
--   지금은 [송장 생성]이 등록 주소를 그대로 써서, 다른 데로 보내려면 고객정보 주소를 고쳤다가
--   되돌려야 했다(고객정보가 오염되고, 어디로 보냈는지 기록도 안 남음).
--
-- 이제 송장 생성 때 이번 건만 쓸 주소를 따로 넣을 수 있고, 그 값을 여기에 남긴다.
--   · 비어 있음  = 고객/거래처 등록 주소로 발송 (기존과 동일 — 기본값)
--   · 값이 있음  = 그 주소로 발송. 화면에 "별도 배송지"로 표시되고, 송장 재발급 때 미리 채워진다
--   · 받는 사람 이름·연락처는 바꾸지 않는다 (사장님 결정) — 고객 식별이 흐려지면 안 되므로
--   · 고객정보(customers)의 주소는 건드리지 않는다

ALTER TABLE offline_sales
  ADD COLUMN IF NOT EXISTS ship_postcode text,
  ADD COLUMN IF NOT EXISTS ship_address_road text,
  ADD COLUMN IF NOT EXISTS ship_address_detail text;

ALTER TABLE deliveries
  ADD COLUMN IF NOT EXISTS ship_postcode text,
  ADD COLUMN IF NOT EXISTS ship_address_road text,
  ADD COLUMN IF NOT EXISTS ship_address_detail text;

COMMENT ON COLUMN offline_sales.ship_address_road IS
  '별도 배송지(도로명). NULL=고객 등록 주소로 발송. 159';
COMMENT ON COLUMN deliveries.ship_address_road IS
  '별도 배송지(도로명). NULL=거래처 등록 주소로 발송. 159';

-- 확인용
SELECT table_name, column_name FROM information_schema.columns
 WHERE column_name LIKE 'ship_address%' AND table_name IN ('offline_sales', 'deliveries')
 ORDER BY table_name, column_name;
