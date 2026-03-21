-- 021: 온라인 주문 재고 차감 추적 플래그
ALTER TABLE orders ADD COLUMN IF NOT EXISTS stock_deducted boolean DEFAULT false;
