import { describe, it, expect } from 'vitest';
import { ctx, jsonReq, idParam, settle } from '../helpers/setup';

/**
 * 판매와 함께 맡은 가위를 한 번에 건네기 (2026-11-22)
 *   실제 사고: 출장에서 제품 판매 + 가위 수리를 함께 받아 손으로 전달 → 판매만 끝나고 수리 건은 「작업완료」에 남음.
 *   따로 누르면 후기 요청이 판매 1통 + 수리 1통으로 두 번 나갈 뻔했다.
 */
describe('판매 [고객 수령 완료] — 함께 맡은 가위도 같이', () => {
  const seed = (repair: Record<string, unknown> = {}, sale: Record<string, unknown> = {}) => {
    ctx.db.seed('offline_sales', [{
      id: 's1', sale_number: 'OS-20261121-001', customer_name: '점검고객', customer_phone: '01000000000',
      invoice_number: null, delivered_at: null, cancelled_at: null,
      review_promised_at: '2026-11-21T00:00:00Z', review_promised_type: 'consult', review_promised_subtype: 'field_request', review_requested_at: null, ...sale,
    }]);
    ctx.db.seed('repairs', [{
      id: 'r1', as_id: 'AS-20261121-0001', sale_id: 's1', status: 'ready_to_ship', name: '점검고객', phone: '01000000000',
      proceed_type: '출장수거', invoice_number: null, qty_mamoru: 2, qty_other: 0,
      review_promised_at: '2026-11-21T00:00:00Z', review_request_sent_at: null, delivered_at: null, delivery_method: null, ...repair,
    }]);
  };
  const pickup = async (body: Record<string, unknown> = {}) => {
    const { PATCH } = await import('@/app/api/sales/[id]/route');
    const res = await PATCH(jsonReq('/api/sales/s1', 'PATCH', { action: 'mark_delivered', mode: 'pickup', ...body }), idParam('s1'));
    await settle();
    return res;
  };

  it('가위도 같이 닫히고, 이력이 남고, 후기 요청은 판매 쪽 1통만', async () => {
    seed();
    const res = await pickup({ repair_ids: ['r1'] });
    const body = await res.json();
    expect(res.status, JSON.stringify(body)).toBe(200);
    expect(body.repairs_handed).toEqual(['AS-20261121-0001']);

    expect(ctx.db.one('offline_sales', 's1')!.delivered_at).toBeTruthy();
    const r = ctx.db.one('repairs', 'r1')!;
    expect(r.status).toBe('delivered');
    expect(r.delivery_method).toBe('pickup');
    expect(r.delivered_at).toBeTruthy();
    expect(r.review_request_sent_at, '3시간 뒤 크론이 한 통 더 보내지 않게 표시').toBeTruthy();

    const h = ctx.db.rows('repair_history');
    expect(h).toHaveLength(1);
    expect([h[0].from_status, h[0].to_status]).toEqual(['ready_to_ship', 'delivered']);

    expect(ctx.sent.filter((s) => s.template === 'as_handed_over'), '가위 전달 안내 1통').toHaveLength(1);
    expect(ctx.sent.filter((s) => s.template !== 'as_handed_over'), '후기 요청은 1통').toHaveLength(1);
    expect(ctx.blocked, '실제 외부 호출이 있으면 안 된다').toEqual([]);
  });

  it('체크를 끄면(가위 id 를 안 보내면) 판매만 끝나고 가위는 그대로', async () => {
    seed();
    const res = await pickup();
    expect(res.status).toBe(200);
    expect(ctx.db.one('repairs', 'r1')!.status).toBe('ready_to_ship');
    expect(ctx.db.rows('repair_history')).toHaveLength(0);
    expect(ctx.sent.map((s) => s.template)).not.toContain('as_handed_over');
  });

  it('아직 작업 전인 가위 · 송장이 있는 가위 · 다른 판매의 가위는 id 를 보내도 건드리지 않는다', async () => {
    seed({ status: 'cost_notified' });
    ctx.db.seed('repairs', [
      { id: 'r2', as_id: 'AS-20261121-0002', sale_id: 's1', status: 'ready_to_ship', name: '점검고객', phone: '01000000000', proceed_type: '출장수거', invoice_number: '000000000000' },
      { id: 'r3', as_id: 'AS-20261121-0003', sale_id: 'other', status: 'ready_to_ship', name: '다른고객', phone: '01000000001', proceed_type: '출장수거', invoice_number: null },
    ]);
    const res = await pickup({ repair_ids: ['r1', 'r2', 'r3'] });
    const body = await res.json();
    expect(res.status).toBe(200);
    expect(body.repairs_handed).toEqual([]);
    expect(ctx.db.one('repairs', 'r1')!.status).toBe('cost_notified');
    expect(ctx.db.one('repairs', 'r2')!.status).toBe('ready_to_ship');
    expect(ctx.db.one('repairs', 'r3')!.status).toBe('ready_to_ship');
    expect(ctx.db.rows('repair_history')).toHaveLength(0);
  });

  it('판매 쪽 후기 요청이 꺼져 있으면 가위 쪽 후기 요청은 살려 둔다(3시간 뒤 크론 몫)', async () => {
    seed({}, { review_promised_at: null });
    const res = await pickup({ repair_ids: ['r1'] });
    expect(res.status).toBe(200);
    const r = ctx.db.one('repairs', 'r1')!;
    expect(r.status).toBe('delivered');
    expect(r.review_request_sent_at).toBeNull();
    expect(ctx.sent.map((s) => s.template)).toEqual(['as_handed_over']);
  });
});

describe('복원수리 [직접 수령] — 판매 쪽에서 후기 요청이 이미 나간 경우', () => {
  const handover = async () => {
    const { PATCH } = await import('@/app/api/repair/[id]/route');
    const res = await PATCH(jsonReq('/api/repair/r1', 'PATCH', { status: 'delivered', delivered_at: '2026-11-22T03:00:00Z', delivery_method: 'pickup', note: '직접 수령 (매장 전달)' }), idParam('r1'));
    await settle();
    return res;
  };
  const seedRepair = () => ctx.db.seed('repairs', [{
    id: 'r1', as_id: 'AS-20261121-0001', sale_id: 's1', status: 'ready_to_ship', name: '점검고객', phone: '01000000000',
    proceed_type: '출장수거', invoice_number: null, paid_at: null, inbound_at: null, qty_mamoru: 2, qty_other: 0,
    review_promised_at: '2026-11-21T00:00:00Z', review_request_sent_at: null,
  }]);

  it('가위 건은 「보낸 것」으로 표시돼 한 통 더 나가지 않는다', async () => {
    ctx.db.seed('offline_sales', [{ id: 's1', sale_number: 'OS-20261121-001', review_requested_at: '2026-11-22T02:42:39Z' }]);
    seedRepair();
    const res = await handover();
    expect(res.status).toBe(200);
    const r = ctx.db.one('repairs', 'r1')!;
    expect(r.status).toBe('delivered');
    expect(r.review_request_sent_at).toBeTruthy();
  });

  it('판매 쪽에서 안 나갔으면 전과 같이 3시간 뒤 크론 몫으로 남는다', async () => {
    ctx.db.seed('offline_sales', [{ id: 's1', sale_number: 'OS-20261121-001', review_requested_at: null }]);
    seedRepair();
    const res = await handover();
    expect(res.status).toBe(200);
    expect(ctx.db.one('repairs', 'r1')!.review_request_sent_at).toBeNull();
  });
});
