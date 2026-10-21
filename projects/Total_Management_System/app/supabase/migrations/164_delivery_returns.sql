-- ═══════════════════════════════════════════════════════════
-- 164: 납품(B2B) 반품 — 부분반품 포함 (2026-10-21)
--   기존: 출고된 납품은 되돌릴 길이 없었다(송장 취소만 가능). returns 는 판매·주문만 연결.
--   이제: 납품 상세 [반품 처리] → returns 건 생성(source='delivery') → **입고 시점에** 한 번만 반영
--         · delivery_items.quantity 차감(+ returned_qty 기록) → 합계 재계산 → 매출은 원래 납품한 달에서 빠짐(사장님 확정)
--         · products.return_stock += 수량 (반품창고 · 판매가능/아임웹 불변 → 재고 화면 [재판매]/[불량])
--   반영 로직 SSOT = app/src/lib/deliveries/apply-return.ts
-- ═══════════════════════════════════════════════════════════

-- returns ↔ 납품 연결 + 반품 품목 목록 (납품은 여러 품목을 한 박스로 회수하므로 행 하나에 items 로 담는다)
--   items: [{ delivery_item_id, product_id, product_name, category, qty, unit_price }]
alter table returns add column if not exists delivery_id uuid references deliveries(id) on delete set null;
alter table returns add column if not exists items jsonb;
-- 입고 반영을 이미 했는지 (멱등 가드 — 같은 반품이 두 번 빠지지 않게)
alter table returns add column if not exists applied_at timestamptz;
create index if not exists idx_returns_delivery on returns(delivery_id);

comment on column returns.delivery_id is '164: 납품(B2B) 반품 연결. source=''delivery''.';
comment on column returns.items       is '164: 납품 반품 품목 [{delivery_item_id, product_id, product_name, category, qty, unit_price}].';
comment on column returns.applied_at  is '164: 납품 반품을 납품 수량·금액·반품창고에 반영한 시각(입고 시 1회).';
comment on column returns.source      is '출처: order(아임웹 주문) | sale(매장판매) | delivery(B2B 납품, 164).';

-- 납품 품목: 반품으로 빠진 수량 (원래 수량 = quantity + returned_qty)
alter table delivery_items add column if not exists returned_qty int not null default 0;

-- 납품: 반품 상태 (null=없음 · 'partial'=일부 반품 · 'full'=전량 반품) + 마지막 반품 반영 시각
alter table deliveries add column if not exists return_status text;
alter table deliveries add column if not exists returned_at timestamptz;

comment on column delivery_items.returned_qty is '164: 반품으로 차감된 수량. 원래 수량 = quantity + returned_qty.';
comment on column deliveries.return_status    is '164: null | partial(일부 반품) | full(전량 반품).';
comment on column deliveries.returned_at      is '164: 마지막 반품 반영(입고) 시각.';
