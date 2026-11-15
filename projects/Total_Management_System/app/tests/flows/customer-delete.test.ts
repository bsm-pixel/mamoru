import { describe, it, expect } from 'vitest';
import { ctx, jsonReq, idParam } from '../helpers/setup';

/** 4차 3번 — 거래·이력이 없는 고객만 삭제된다 */
describe('고객 삭제', () => {
  it('거래가 없으면 메모와 함께 삭제된다', async () => {
    ctx.db.seed('customers', [{ id: 'c1', name: '빈 고객' }]);
    ctx.db.seed('customer_notes', [{ id: 'n1', customer_id: 'c1', note: '메모' }]);
    const { DELETE } = await import('@/app/api/customers/[id]/route');
    const res = await DELETE(jsonReq('/api/customers/c1', 'DELETE'), idParam('c1'));
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    expect(ctx.db.rows('customers')).toHaveLength(0);
    expect(ctx.db.rows('customer_notes')).toHaveLength(0);
  });
  it('취소된 판매 1건만 있어도 막고, 무엇이 걸렸는지 알려 준다', async () => {
    ctx.db.seed('customers', [{ id: 'c1', name: '거래 고객' }]);
    ctx.db.seed('offline_sales', [{ id: 's1', customer_id: 'c1', cancelled_at: '2026-11-15T00:00:00Z' }]);
    ctx.db.seed('repairs', [{ id: 'r1', customer_id: 'c1' }, { id: 'r2', customer_id: 'c1' }]);
    const { DELETE } = await import('@/app/api/customers/[id]/route');
    const res = await DELETE(jsonReq('/api/customers/c1', 'DELETE'), idParam('c1'));
    const body = await res.json();
    expect(res.status).toBe(400);
    expect(body.error).toContain('판매 1건');
    expect(body.error).toContain('복원수리 2건');
    expect(ctx.db.rows('customers')).toHaveLength(1);
  });
  it('다른 고객이 병합돼 있으면 막는다 · 로그인 없으면 거부', async () => {
    ctx.db.seed('customers', [{ id: 'c1', name: '대표' }, { id: 'c2', name: '흡수됨', merged_into_id: 'c1' }]);
    const { DELETE } = await import('@/app/api/customers/[id]/route');
    const res = await DELETE(jsonReq('/api/customers/c1', 'DELETE'), idParam('c1'));
    expect(res.status).toBe(400);
    expect((await res.json()).error).toContain('병합된 고객 1건');
    ctx.db.user = null;
    const res2 = await DELETE(jsonReq('/api/customers/c2', 'DELETE'), idParam('c2'));
    expect(res2.status).toBe(401);
    expect(ctx.db.rows('customers')).toHaveLength(2);
  });
});
