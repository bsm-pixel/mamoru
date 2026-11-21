import { describe, it, expect } from 'vitest';
import { ctx, jsonReq, idParam, settle } from '../helpers/setup';
import { getAllowedRepairTransitions } from '@/lib/repair/transitions';

/** ⑤ 복원수리 접수 상태 변경 — 알림톡은 실제로 보내지 않고 "보내려 했다"만 확인 */
describe('복원수리 상태 변경', () => {
  const seed = (status: string) => ctx.db.seed('repairs', [{
    id: 'r1', as_id: 'AS-20261115-0001', status, name: '점검고객', phone: '01000000000',
    proceed_type: '택배수거', invoice_number: null, paid_at: null, inbound_at: null,
  }]);

  it('허용된 다음 단계로 바꾸면 이력이 남고, 외부로는 아무것도 나가지 않는다', async () => {
    seed('cost_notified');
    const next = getAllowedRepairTransitions('cost_notified' as never)[0];
    expect(next, '비용 안내 다음 단계가 정의돼 있어야 한다').toBeTruthy();
    const { PATCH } = await import('@/app/api/repair/[id]/route');
    const res = await PATCH(jsonReq('/api/repair/r1', 'PATCH', { status: next, note: '자동 점검' }), idParam('r1'));
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    await settle();
    expect(ctx.db.one('repairs', 'r1')!.status).toBe(next);
    const h = ctx.db.rows('repair_history');
    expect(h).toHaveLength(1);
    expect([h[0].from_status, h[0].to_status]).toEqual(['cost_notified', next]);
    expect(ctx.blocked, '실제 외부 호출이 있으면 안 된다').toEqual([]);
  });

  it('건너뛰는 변경은 막고 상태·이력은 그대로다', async () => {
    seed('cost_notified');
    const allowed = getAllowedRepairTransitions('cost_notified' as never) as string[];
    const bad = ['completed', 'delivered', 'shipped'].find((s) => !allowed.includes(s))!;
    const { PATCH } = await import('@/app/api/repair/[id]/route');
    const res = await PATCH(jsonReq('/api/repair/r1', 'PATCH', { status: bad }), idParam('r1'));
    expect(res.status).toBe(400);
    expect(ctx.db.one('repairs', 'r1')!.status).toBe('cost_notified');
    expect(ctx.db.rows('repair_history')).toHaveLength(0);
    expect(ctx.sent).toHaveLength(0);
  });
});

/** 직접 전달(2026-11-21) — 택배 건을 송장 없이 손으로 건넨 경우 */
describe('복원수리 직접 전달', () => {
  const seed = (extra: Record<string, unknown> = {}) => ctx.db.seed('repairs', [{
    id: 'r2', as_id: 'AS-20261121-0002', status: 'repairing', name: '점검고객', phone: '01000000000',
    proceed_type: '출장수거', invoice_number: null, paid_at: null, inbound_at: null,
    qty_mamoru: 2, qty_other: 1, review_promised_at: '2026-11-21T00:00:00Z', review_request_sent_at: null, ...extra,
  }]);
  const handover = async () => {
    const { PATCH } = await import('@/app/api/repair/[id]/route');
    const res = await PATCH(jsonReq('/api/repair/r2', 'PATCH', { status: 'delivered', delivered_at: '2026-11-21T10:00:00Z', delivery_method: 'pickup', note: '직접 수령 (매장 전달)' }), idParam('r2'));
    await settle();
    return res;
  };

  it('전달 안내(수리내역 조회)만 나가고, 후기 요청은 바로 나가지 않는다', async () => {
    seed();
    const res = await handover();
    expect(res.status).toBe(200);
    expect(ctx.db.one('repairs', 'r2')!.status).toBe('delivered');
    expect(ctx.sent.map((s) => s.template)).toEqual(['as_handed_over']);
    expect(ctx.sent[0].data).toMatchObject({ as_id: 'AS-20261121-0002', as_uid: 'AS-20261121-0002', qty: '3' });
    expect(ctx.db.one('repairs', 'r2')!.review_request_sent_at, '후기 요청은 3시간 뒤 크론 몫').toBeNull();
  });

  it('송장이 있는 건의 배송완료는 전과 같이 후기 요청이 바로 나간다', async () => {
    ctx.db.seed('repairs', [{
      id: 'r3', as_id: 'AS-20261121-0003', status: 'shipped', name: '점검고객', phone: '01000000000',
      proceed_type: '직접발송', invoice_number: '000000000000', review_promised_at: '2026-11-21T00:00:00Z', review_request_sent_at: null,
    }]);
    const { PATCH } = await import('@/app/api/repair/[id]/route');
    const res = await PATCH(jsonReq('/api/repair/r3', 'PATCH', { status: 'delivered' }), idParam('r3'));
    await settle();
    expect(res.status).toBe(200);
    expect(ctx.sent.map((s) => s.template)).not.toContain('as_handed_over');
  });
});
