-- 105: EVENT 캠페인 수량 묶음 할인 규칙
-- discount_rules: [{ unit_price, min_qty, bundle_price }]
--   같은 단가끼리 묶음, min_qty 도달 시 묶음 반복 + 나머지 정가.
--   예: {unit_price:50000, min_qty:3, bundle_price:130000} → 5만원 3자루=13만, 4자루=18만, 6자루=26만
alter table event_campaigns add column if not exists discount_rules jsonb not null default '[]'::jsonb;
