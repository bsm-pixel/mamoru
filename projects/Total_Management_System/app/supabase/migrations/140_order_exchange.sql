-- ═══════════════════════════════════════════════════════════
-- 140: 주문(아임웹 온라인) 제품 교환 — 이력/감사 컬럼
--   교환 시 아임웹 주문·카드결제·매출은 '불변', 상품/재고만 스왑한다.
--   (반납품 → 반품창고, 새 제품 → 시리얼 출고, 차액 → cash_transactions '주문교환차액')
-- ═══════════════════════════════════════════════════════════
alter table orders add column if not exists exchanged_at  timestamptz;
alter table orders add column if not exists exchange_memo text;

comment on column orders.exchanged_at  is '제품 교환 처리 시각. 있으면 이 주문은 교환됨(원 결제/매출은 유지, 상품만 바뀜).';
comment on column orders.exchange_memo is '교환 내역 요약(반납/발송 품목·회수방식·차액).';
