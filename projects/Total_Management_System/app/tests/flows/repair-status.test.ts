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
