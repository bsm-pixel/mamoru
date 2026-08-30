-- ═══════════════════════════════════════════════════════════
-- 141: 주문 교환 — 새 제품 발송 방식 + 교환품 송장
--   발송 방식(배송/직접전달) 선택, 배송이면 새 제품용 롯데 송장을 별도 발행.
--   (원 주문 invoice_number/status 는 원본 배송 그대로 — 교환품은 별도 송장으로 추적)
-- ═══════════════════════════════════════════════════════════
alter table orders add column if not exists exchange_ship_method    text;        -- '배송' | '직접전달'
alter table orders add column if not exists exchange_goods          text;        -- 교환 새 제품명(송장 품목명)
alter table orders add column if not exists exchange_invoice_number text;        -- 교환품 발송 송장번호(롯데)
alter table orders add column if not exists exchange_shipped_at      timestamptz; -- 교환품 송장 발행 시각

comment on column orders.exchange_ship_method    is '교환 새 제품 발송 방식: 배송(송장 필요) / 직접전달(송장 불필요)';
comment on column orders.exchange_invoice_number is '교환품 발송 롯데 송장번호(원 주문 송장과 별개).';
