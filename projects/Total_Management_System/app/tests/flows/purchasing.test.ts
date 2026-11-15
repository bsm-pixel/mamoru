import { describe, it, expect } from 'vitest';
import { ctx, jsonReq, idParam } from '../helpers/setup';

/** ③ 발주 생성 → 취소 → 삭제 */
describe('발주 생성 → 취소 → 삭제', () => {
  it('금액이 맞게 만들어지고, 취소해야만 지워지며, 재고는 변하지 않는다', async () => {
    ctx.db.seed('products', [{ id: 'p1', name: '블런트 5.5', stock_quantity: 5, raw_stock: 5 }]);
    const { POST } = await import('@/app/api/purchasing/route');
    const res = await POST(jsonReq('/api/purchasing', 'POST', {
      supplier_name: '점검 공급처', order_date: '2026-11-15', vat_type: 'included',
      items: [{ product_id: 'p1', product_name: '블런트 5.5', quantity: 10, unit_price: 11000 }],
    }));
    const body = await res.json();
    expect(res.status, JSON.stringify(body)).toBe(200);
    const po = ctx.db.rows('purchase_orders')[0];
    expect(po.po_number).toBe('PO-20261115-001');
    expect(po.total_amount).toBe(110000);
    expect(po.supply_amount).toBe(100000);
    expect(po.vat_amount).toBe(10000);
    expect(ctx.db.rows('purchase_order_items')).toHaveLength(1);

    const { PATCH, DELETE } = await import('@/app/api/purchasing/[id]/route');
    // 취소 전에는 삭제 불가
    const early = await DELETE(jsonReq(`/api/purchasing/${po.id}`, 'DELETE'), idParam(po.id));
    expect(early.status).toBe(400);
    expect(ctx.db.rows('purchase_orders')).toHaveLength(1);

    const cancel = await PATCH(jsonReq(`/api/purchasing/${po.id}`, 'PATCH', { status: 'cancelled' }), idParam(po.id));
    expect(cancel.status, JSON.stringify(await cancel.clone().json())).toBe(200);
    expect(ctx.db.one('purchase_orders', po.id)!.status).toBe('cancelled');

    const del = await DELETE(jsonReq(`/api/purchasing/${po.id}`, 'DELETE'), idParam(po.id));
    expect(del.status).toBe(200);
    expect(ctx.db.rows('purchase_orders')).toHaveLength(0);
    expect(ctx.db.rows('purchase_order_items')).toHaveLength(0);
    expect(ctx.db.one('products', 'p1')!.stock_quantity).toBe(5); // 입고 전이므로 재고 불변
  });
});
