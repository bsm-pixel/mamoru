import { describe, it, expect } from 'vitest';
import { ctx, jsonReq, idParam } from '../helpers/setup';

/** ④ 고객 등록 · 수정 · 병합 */
describe('고객 등록 · 수정 · 병합', () => {
  it('등록하고, 같은 전화번호는 막고, 수정이 반영된다', async () => {
    const { POST } = await import('@/app/api/customers/route');
    const a = await POST(jsonReq('/api/customers', 'POST', { name: '점검고객', phone: '010-1111-2222' }));
    const aBody = await a.json();
    expect(a.status, JSON.stringify(aBody)).toBe(200);
    const id = aBody.customer.id;
    expect(ctx.db.one('customers', id)!.customer_type).toBe('retail');

    const dup = await POST(jsonReq('/api/customers', 'POST', { name: '다른이름', phone: '01011112222' }));
    expect(dup.status).toBe(409);
    expect(ctx.db.rows('customers')).toHaveLength(1);

    const { PATCH } = await import('@/app/api/customers/[id]/route');
    const up = await PATCH(jsonReq(`/api/customers/${id}`, 'PATCH', { name: '점검고객2', memo: '수정됨' }), idParam(id));
    expect(up.status, JSON.stringify(await up.clone().json())).toBe(200);
    expect(ctx.db.one('customers', id)!.name).toBe('점검고객2');
    expect(ctx.db.one('customers', id)!.memo).toBe('수정됨');
  });

  it('병합은 데이터베이스 함수에 대표·대상 고객을 정확히 넘긴다', async () => {
    ctx.db.rpcHandlers.merge_customers = () => ({ merged: 1 });
    const { POST } = await import('@/app/api/customers/merge/route');
    const res = await POST(jsonReq('/api/customers/merge', 'POST', { primaryId: 'c1', victimIds: ['c2'] }));
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    expect(ctx.db.rpcCalls).toEqual([{ name: 'merge_customers', args: { p_primary: 'c1', p_victims: ['c2'] } }]);
  });

  it('자기 자신과의 병합·빈 대상은 거부한다', async () => {
    const { POST } = await import('@/app/api/customers/merge/route');
    const res = await POST(jsonReq('/api/customers/merge', 'POST', { primaryId: 'c1', victimIds: [] }));
    expect(res.status).toBe(400);
    expect(ctx.db.rpcCalls).toHaveLength(0);
  });
});
