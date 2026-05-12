-- 081: purchase_order_items.received_quantity — 입고검수 시 실수령 수량 (제작품이라 주문≠입고 발생)
--
-- 배경 (2026-05-13):
--   가위는 제작품이라 90자루 주문해도 85자루만 만들어져 오는 경우가 흔함. 기존엔 "입고 확인" 시
--   주문 수량(quantity)만큼 무조건 재고를 늘리고 금액도 그대로였음 → 실수령 수량을 따로 기록하고
--   재고·금액(잔금)을 실수령 기준으로 보정할 수 있도록 컬럼 추가.
--   NULL = 아직 입고 전 (입고 시 quantity 그대로 또는 조정값으로 채워짐). quantity(주문 수량)는 절대 덮어쓰지 않음 — 발주 기록 보존.
ALTER TABLE purchase_order_items ADD COLUMN IF NOT EXISTS received_quantity integer;
COMMENT ON COLUMN purchase_order_items.received_quantity IS '입고검수 시 실수령 수량. NULL = 입고 전. quantity(주문)는 그대로 둠';
