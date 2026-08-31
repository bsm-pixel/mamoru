-- ═══════════════════════════════════════════════════════════
-- 142: returns 를 '교환·반품 SSOT'로 확장 — 주문(아임웹) 교환도 여기서 관리
--   기존 returns 는 sale_id(매장판매)만 연결 → 주문 교환도 추적하도록 order_id·source 추가.
--   교환 상세(발송품·송장·발송방식)는 주문 교환건은 orders 의 exchange_* 컬럼(140·141)을 참조.
-- ═══════════════════════════════════════════════════════════
alter table returns add column if not exists order_id uuid references orders(id) on delete set null;
alter table returns add column if not exists source   text default 'sale';   -- 'order'(아임웹 주문) | 'sale'(매장판매)

create index if not exists idx_returns_order on returns(order_id);

comment on column returns.order_id is '아임웹 주문 교환건 연결(주문 상세는 배지+링크, 상세는 여기서 관리).';
comment on column returns.source   is '출처: order(아임웹 주문) | sale(매장판매).';
