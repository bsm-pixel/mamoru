import { describe, it, expect, vi } from 'vitest';
import { ctx, jsonReq, idParam, settle } from '../helpers/setup';

/**
 * 출장 상담 주소 (2026-11-22, 작업실 t-261121-field-address · 사장님 지시)
 *  - TMS 에서 주소를 고치면 좌표를 다시 잡고 이력을 남기며, 알림톡은 보내지 않는다
 *  - 고객 정보 주소는 고를 때만 바뀐다
 *  - 출장 접수는 기존 고객의 저장된 주소를 덮어쓰지 않는다(다르면 상세에서 확인 창). 새 고객은 접수 주소가 그대로 들어간다
 */
vi.mock('@/lib/kakao/geocode', () => ({ geocodeForConsultation: vi.fn(async (road: string) => (road.includes('못찾는') ? null : { lat: 37.1, lng: 126.9 })) }));
vi.mock('@/lib/google/calendar-sync', () => ({ syncConsultationToCalendar: vi.fn(async () => {}) }));
vi.mock('@/lib/notification/email', () => ({ sendAdminEmail: vi.fn(async () => true) }));

const seed = () => {
  ctx.db.seed('customers', [{ id: 'cu1', name: '점검고객', phone: '01000000009', phone_normalized: '01000000009', address_road: '인천 미추홀구 옛길 1', address_detail: '101호', postcode: '11111', merged_into_id: null, created_at: '2026-01-01T00:00:00Z' }]);
  ctx.db.seed('consultations', [{
    id: 'k1', unique_id: 'u1', name: '점검고객', phone: '01000000009', consultation_type: 'field_request', status: 'pending_admin',
    customer_id: 'cu1', address_road: '인천 미추홀구 틀린길 9', address_detail: '9층', postcode: '22222',
    latitude: 1, longitude: 1, address_sido: '인천', address_sigungu: '미추홀구', address_region: '구', gas_raw: { days: '월' }, visit_date: null, visit_time: null, admin_note: null,
  }]);
};
const patch = async (body: Record<string, unknown>) => {
  const { PATCH } = await import('@/app/api/consultation/[id]/route');
  const res = await PATCH(jsonReq('/api/consultation/k1', 'PATCH', body), idParam('k1'));
  await settle();
  return res;
};

describe('출장 상담 주소 고치기', () => {
  it('이 건만 고치면: 주소 · 좌표가 바뀌고 이력이 남고, 고객 주소와 알림톡은 그대로', async () => {
    seed();
    const res = await patch({ postcode: '33333', address_road: '인천 남동구 새길 5', address_detail: '2층', also_customer: false });
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    const k = ctx.db.one('consultations', 'k1')!;
    expect([k.address_road, k.address_detail, k.postcode]).toEqual(['인천 남동구 새길 5', '2층', '33333']);
    expect([k.latitude, k.longitude]).toEqual([37.1, 126.9]);
    expect([k.address_sido, k.address_sigungu, k.address_region]).toEqual([null, null, null]);
    expect(k.gas_raw.days, '접수 때 값은 남는다').toBe('월');
    expect(k.gas_raw.address_decision).toBe('kept');
    expect(ctx.db.one('customers', 'cu1')!.address_road).toBe('인천 미추홀구 옛길 1');
    const h = ctx.db.rows('consultation_history');
    expect(h).toHaveLength(1);
    expect(h[0].from_status).toBe(h[0].to_status);
    expect(h[0].note).toContain('인천 미추홀구 틀린길 9');
    expect(ctx.sent, '알림톡은 나가지 않는다').toEqual([]);
  });

  it('고객 주소도 바꾸면: 고객 정보 주소까지 바뀐다', async () => {
    seed();
    await patch({ postcode: '33333', address_road: '인천 남동구 새길 5', address_detail: '2층', also_customer: true });
    const c = ctx.db.one('customers', 'cu1')!;
    expect([c.address_road, c.address_detail, c.postcode]).toEqual(['인천 남동구 새길 5', '2층', '33333']);
    expect(ctx.db.one('consultations', 'k1')!.gas_raw.address_decision).toBe('applied');
  });

  it('새 주소의 지도 위치를 못 찾으면 옛 좌표 대신 비워 둔다', async () => {
    seed();
    await patch({ address_road: '못찾는 주소', address_detail: '' });
    const k = ctx.db.one('consultations', 'k1')!;
    expect([k.latitude, k.longitude]).toEqual([null, null]);
  });

  it('도로명 주소를 비우면 막는다', async () => {
    seed();
    const res = await patch({ address_road: '  ', address_detail: '2층' });
    expect(res.status).toBe(400);
    expect(ctx.db.one('consultations', 'k1')!.address_road).toBe('인천 미추홀구 틀린길 9');
  });
});

describe('접수 주소가 고객 정보 주소와 다를 때 — 확인 창에서 고른 것', () => {
  it('[고객 주소로 바꾸기] — 고객 정보 주소가 접수 주소로 바뀌고, 상담 주소는 그대로', async () => {
    seed();
    const res = await patch({ address_to_customer: 'apply' });
    expect(res.status).toBe(200);
    const c = ctx.db.one('customers', 'cu1')!;
    expect([c.address_road, c.address_detail, c.postcode]).toEqual(['인천 미추홀구 틀린길 9', '9층', '22222']);
    const k = ctx.db.one('consultations', 'k1')!;
    expect(k.address_road).toBe('인천 미추홀구 틀린길 9');
    expect(k.gas_raw.address_decision).toBe('applied');
    expect(ctx.sent).toEqual([]);
  });

  it('[그대로 두기] — 고객 정보 주소는 그대로, 다시 묻지 않게 기록만', async () => {
    seed();
    await patch({ address_to_customer: 'keep' });
    expect(ctx.db.one('customers', 'cu1')!.address_road).toBe('인천 미추홀구 옛길 1');
    expect(ctx.db.one('consultations', 'k1')!.gas_raw.address_decision).toBe('kept');
    expect(ctx.db.rows('consultation_history')).toHaveLength(0);
  });
});

describe('출장 접수 — 고객 주소', () => {
  const submit = async (phone: string) => {
    const { POST } = await import('@/app/api/consultation/public/submit/route');
    const res = await POST(jsonReq('/api/consultation/public/submit', 'POST', {
      name: '점검고객', phone, type: '출장 요청', addressRoad: '인천 남동구 접수길 3', addressDetail: '3층', addressZip: '44444', days: ['월'], timePrefs: ['오전'],
    }));
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    await settle();
  };

  it('기존 고객: 저장된 주소를 덮어쓰지 않는다(상세에서 확인 창으로 고른다)', async () => {
    seed();
    ctx.db.rows('consultations').length = 0;
    await submit('010-0000-0009');
    expect(ctx.db.one('customers', 'cu1')!.address_road).toBe('인천 미추홀구 옛길 1');
    const k = ctx.db.rows('consultations')[0];
    expect(k.address_road, '상담 건에는 접수 주소가 들어간다').toBe('인천 남동구 접수길 3');
  });

  it('기존 고객인데 저장된 주소가 비어 있으면: 접수 주소로 채운다', async () => {
    ctx.db.seed('customers', [{ id: 'cu2', name: '점검고객', phone: '01000000008', phone_normalized: '01000000008', address_road: null, address_detail: null, postcode: null, merged_into_id: null, created_at: '2026-01-01T00:00:00Z' }]);
    await submit('010-0000-0008');
    expect(ctx.db.one('customers', 'cu2')!.address_road).toBe('인천 남동구 접수길 3');
  });

  it('새 고객: 접수 주소가 그대로 고객 주소가 된다', async () => {
    await submit('010-0000-0007');
    const c = ctx.db.rows('customers').find((x) => x.phone === '010-0000-0007');
    expect(c?.address_road).toBe('인천 남동구 접수길 3');
    expect(c?.address_detail).toBe('3층');
  });
});
