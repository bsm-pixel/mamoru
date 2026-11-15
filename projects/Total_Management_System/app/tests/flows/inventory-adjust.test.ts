import { describe, it, expect } from 'vitest';
import { ctx, jsonReq } from '../helpers/setup';

/** ② 재고 조정 +/- */
describe('재고 조정', () => {
  const seed = () => ctx.db.seed('products', [{ id: 'p1', name: '틴닝 6.0', stock_quantity: 4, raw_stock: 4, imweb_product_no: null }]);

  it('늘리고 줄이면 재고와 이력이 맞는다', async () => {
    seed();
    const { POST } = await import('@/app/api/inventory/adjust/route');
    const up = await POST(jsonReq('/api/inventory/adjust', 'POST', { product_id: 'p1', adjustment_type: 'correction', quantity: 3, reason: '점검' }));
    expect(up.status, JSON.stringify(await up.clone().json())).toBe(200);
    expect(ctx.db.one('products', 'p1')!.stock_quantity).toBe(7);
    const down = await POST(jsonReq('/api/inventory/adjust', 'POST', { product_id: 'p1', adjustment_type: 'damage', quantity: -5, reason: '점검' }));
    expect(down.status).toBe(200);
    expect(ctx.db.one('products', 'p1')!.stock_quantity).toBe(2);
    expect(ctx.db.one('products', 'p1')!.raw_stock).toBe(2);
    expect(ctx.db.rows('stock_adjustments').map((a) => a.quantity)).toEqual([3, -5]);
  });

  it('재고보다 많이 줄이려 하면 막고 재고는 그대로다', async () => {
    seed();
    const { POST } = await import('@/app/api/inventory/adjust/route');
    const res = await POST(jsonReq('/api/inventory/adjust', 'POST', { product_id: 'p1', adjustment_type: 'damage', quantity: -9 }));
    expect(res.status).toBe(400);
    expect(ctx.db.one('products', 'p1')!.stock_quantity).toBe(4);
    expect(ctx.db.rows('stock_adjustments')).toHaveLength(0);
  });

  it('로그인 없이는 거부한다', async () => {
    seed(); ctx.db.user = null;
    const { POST } = await import('@/app/api/inventory/adjust/route');
    const res = await POST(jsonReq('/api/inventory/adjust', 'POST', { product_id: 'p1', adjustment_type: 'correction', quantity: 1 }));
    expect(res.status).toBe(401);
    expect(ctx.db.one('products', 'p1')!.stock_quantity).toBe(4);
  });
});
