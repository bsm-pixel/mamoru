-- 016: 재고 관리 강화 — 창고 구분 + 미입고 수량 뷰

-- 시리얼별 창고 구분 (보관/진열)
ALTER TABLE product_serials ADD COLUMN IF NOT EXISTS warehouse_zone text DEFAULT 'storage';
-- storage(보관) | display(진열)

-- 미입고 수량 뷰 (발주완료/선납완료 상태의 PO 품목 합계)
CREATE OR REPLACE VIEW v_pending_stock AS
SELECT poi.product_id, SUM(poi.quantity) as pending_qty
FROM purchase_order_items poi
JOIN purchase_orders po ON po.id = poi.po_id
WHERE po.status IN ('ordered','deposit_paid')
GROUP BY poi.product_id;
