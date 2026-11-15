import { describe, it, expect } from 'vitest';
import { ctx, jsonReq, idParam } from '../helpers/setup';

/** 4차 1번 — 이력이 있는 제품 삭제는 사람 말로 막고, 이력이 없으면 지워진다 */
describe('제품 삭제', () => {
  it('재고 조정 이력이 있으면 400과 읽을 수 있는 문장, 제품은 그대로', async () => {
    ctx.db.seed('products', [{ id: 'p1', name: '점검 제품' }]);
    ctx.db.seed('stock_adjustments', [{ id: 'a1', product_id: 'p1', quantity: 5 }, { id: 'a2', product_id: 'p1', quantity: -5 }]);
    const { DELETE } = await import('@/app/api/products/[id]/route');
    const res = await DELETE(jsonReq('/api/products/p1', 'DELETE'), idParam('p1'));
    const body = await res.json();
    expect(res.status).toBe(400);
    expect(typeof body.error).toBe('string');
    expect(body.error).toContain('재고 조정 2건');
    expect(body.error).not.toContain('[object');
    expect(ctx.db.rows('products')).toHaveLength(1);
  });
  it('이력이 없으면 삭제된다', async () => {
    ctx.db.seed('products', [{ id: 'p2', name: '빈 제품' }]);
    const { DELETE } = await import('@/app/api/products/[id]/route');
    const res = await DELETE(jsonReq('/api/products/p2', 'DELETE'), idParam('p2'));
    expect(res.status).toBe(200);
    expect(ctx.db.rows('products')).toHaveLength(0);
  });
});
