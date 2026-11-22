import { describe, it, expect } from 'vitest';
import { ctx, jsonReq, idParam, settle } from '../helpers/setup';
import { maskName, signMessage, signUrl, cleanEngraving, isValidSignature } from '@/lib/contracts/sign-link';

/**
 * 계약서 고객 서명 링크 + 이니셜 각인 + 판매 연결 (2026-11-22, 작업실 t-261122-econtract-link, 마이그 172)
 *  - 링크는 한 번 만들면 같은 주소 · 고객 화면엔 전화 · 주소가 나가지 않는다
 *  - 서명은 한 번만 · 동의 없이는 안 된다 · 서명되면 사장님 푸시(알림톡은 안 나감)
 *  - 각인 문구 「나중에 받기」 → 같은 링크로 한 번 받는다
 *  - 계약서로 판매를 만들면 계약서에 판매가 이어지고, 판매를 취소하면 풀린다
 */
const SIG = 'data:image/png;base64,' + 'A'.repeat(400);
const T = '11111111-2222-4333-8444-555555555555';

const seed = (over: Record<string, unknown> = {}) => {
  ctx.db.seed('contracts', [{
    id: 'ct1', contract_number: 'CT-20261122-001', customer_name: '김점검', customer_phone: '01000000001', customer_address: '서울 비밀주소 1',
    status: 'draft', total_amount: 500000, discount_amount: 20000, final_amount: 480000, payment_method: 'transfer', installment_months: 0,
    deposit_amount: 0, balance_amount: 480000, signature_data: null, signed_at: null, offline_sale_id: null, created_at: '2026-11-22T01:00:00Z',
    sign_token: null, engraving: false, engraving_text: null, engraving_status: null, ...over,
  }]);
  ctx.db.seed('contract_items', [{ id: 'ci1', contract_id: 'ct1', product_id: 'p1', product_name: '블런트 6.0', quantity: 1, unit_price: 500000, total_price: 500000 }]);
};
const pub = async () => (await import('@/app/api/contracts/public/route'));
const get = async (t: string) => { const { GET } = await pub(); return GET(jsonReq(`/api/contracts/public?t=${t}`, 'GET')); };
const post = async (body: Record<string, unknown>) => { const { POST } = await pub(); const r = await POST(jsonReq('/api/contracts/public', 'POST', body)); await settle(); return r; };

describe('서명 링크 만들기', () => {
  it('처음 누르면 열쇠가 생기고, 다시 눌러도 같은 주소', async () => {
    seed();
    const { POST } = await import('@/app/api/contracts/[id]/sign-link/route');
    const a = await (await POST(jsonReq('/api/contracts/ct1/sign-link', 'POST'), idParam('ct1'))).json();
    const b = await (await POST(jsonReq('/api/contracts/ct1/sign-link', 'POST'), idParam('ct1'))).json();
    expect(a.token).toMatch(/^[0-9a-f-]{36}$/);
    expect(b.token).toBe(a.token);
    expect(a.url).toBe(signUrl(a.token));
    expect(ctx.db.one('contracts', 'ct1')!.sign_link_at).toBeTruthy();
  });

  it('취소된 계약서는 링크를 만들지 않는다', async () => {
    seed({ status: 'cancelled' });
    const { POST } = await import('@/app/api/contracts/[id]/sign-link/route');
    expect((await POST(jsonReq('/api/contracts/ct1/sign-link', 'POST'), idParam('ct1'))).status).toBe(400);
  });
});

describe('고객 화면 — 계약 내용 보기', () => {
  it('내용 · 금액 · 문구는 보이고, 전화 · 주소는 나가지 않으며 이름은 가린다', async () => {
    seed({ sign_token: T });
    const res = await get(T);
    const body = await res.json();
    expect(res.status).toBe(200);
    expect(body.data.name).toBe('김*검');
    expect(body.data.final).toBe(480000);
    expect(body.data.items).toEqual([{ name: '블런트 6.0', qty: 1, price: 500000, total: 500000 }]);
    expect(body.data.notices.caution).toContain('각인');
    const raw = JSON.stringify(body);
    expect(raw).not.toContain('01000000001');
    expect(raw).not.toContain('비밀주소');
    expect(res.headers.get('access-control-allow-origin')).toBe('*');
  });

  it('열쇠가 틀리거나 모양이 아니면 못 연다', async () => {
    seed({ sign_token: T });
    expect((await get('99999999-2222-4333-8444-555555555555')).status).toBe(404);
    expect((await get('abc')).status).toBe(400);
  });
});

describe('고객 서명', () => {
  it('동의 + 서명 → 서명 완료 · 링크 서명 표시 · 사장님 푸시 1건 · 알림톡 0', async () => {
    seed({ sign_token: T });
    const res = await post({ t: T, action: 'sign', signature: SIG, agreed: true });
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    const c = ctx.db.one('contracts', 'ct1')!;
    expect(c.status).toBe('signed');
    expect(c.signature_data).toBe(SIG);
    expect(c.signed_via).toBe('link');
    expect(c.terms_agreed_at).toBeTruthy();
    expect(ctx.pushes).toHaveLength(1);
    expect(ctx.pushes[0].url).toBe('/contracts/ct1');
    expect(ctx.sent).toEqual([]);
  });

  it('동의 안 함 · 서명 없음 · 두 번째 서명은 막는다', async () => {
    seed({ sign_token: T });
    expect((await post({ t: T, action: 'sign', signature: SIG, agreed: false })).status).toBe(400);
    expect((await post({ t: T, action: 'sign', signature: 'data:image/png;base64,', agreed: true })).status).toBe(400);
    expect((await post({ t: T, action: 'sign', signature: SIG, agreed: true })).status).toBe(200);
    expect((await post({ t: T, action: 'sign', signature: SIG.replace('A', 'B'), agreed: true })).status).toBe(409);
    expect(ctx.db.one('contracts', 'ct1')!.signature_data).toBe(SIG);
  });

  it('취소된 계약서에는 서명할 수 없다', async () => {
    seed({ sign_token: T, status: 'cancelled' });
    expect((await post({ t: T, action: 'sign', signature: SIG, agreed: true })).status).toBe(409);
  });
});

describe('이니셜 각인', () => {
  it('각인 계약은 각인 동의까지 있어야 서명된다', async () => {
    seed({ sign_token: T, engraving: true, engraving_status: 'pending' });
    expect((await post({ t: T, action: 'sign', signature: SIG, agreed: true })).status).toBe(400);
    const ok = await post({ t: T, action: 'sign', signature: SIG, agreed: true, engraving_agreed: true });
    expect((await ok.json()).engravingPending).toBe(true);
  });

  it('서명하면서 문구를 적으면 바로 받음', async () => {
    seed({ sign_token: T, engraving: true, engraving_status: 'pending' });
    await post({ t: T, action: 'sign', signature: SIG, agreed: true, engraving_agreed: true, engraving_text: '  S.M.B  ' });
    const c = ctx.db.one('contracts', 'ct1')!;
    expect([c.engraving_text, c.engraving_status]).toEqual(['S.M.B', 'received']);
  });

  it('나중에 받기 → 같은 링크로 한 번 보내면 받음 · 두 번째는 막음 · 푸시', async () => {
    seed({ sign_token: T, engraving: true, engraving_status: 'pending', signature_data: SIG, status: 'signed' });
    expect((await post({ t: T, action: 'engraving', text: '' })).status).toBe(400);
    const r = await post({ t: T, action: 'engraving', text: 'K.J' });
    expect(r.status).toBe(200);
    expect(ctx.db.one('contracts', 'ct1')!.engraving_status).toBe('received');
    expect((await post({ t: T, action: 'engraving', text: 'X' })).status).toBe(409);
    expect(ctx.db.one('contracts', 'ct1')!.engraving_text).toBe('K.J');
    expect(ctx.pushes.some((p) => String(p.body).includes('K.J'))).toBe(true);
  });

  it('직원이 상세에서 받은 문구를 적고, 비우면 다시 기다림', async () => {
    seed({ engraving: true, engraving_status: 'pending' });
    const { PATCH } = await import('@/app/api/contracts/[id]/route');
    await PATCH(jsonReq('/api/contracts/ct1', 'PATCH', { engraving_text: 'A.B' }), idParam('ct1'));
    expect(ctx.db.one('contracts', 'ct1')!.engraving_status).toBe('received');
    await PATCH(jsonReq('/api/contracts/ct1', 'PATCH', { engraving_text: '' }), idParam('ct1'));
    const c = ctx.db.one('contracts', 'ct1')!;
    expect([c.engraving_text, c.engraving_status]).toEqual([null, 'pending']);
  });

  it('계약서 저장 때 각인 문구가 없으면 「나중에 받기」로', async () => {
    const { POST } = await import('@/app/api/contracts/route');
    const res = await POST(jsonReq('/api/contracts', 'POST', {
      contract: { customer_id: 'cu9', customer_name: '이점검', total_amount: 100, final_amount: 100, payment_method: 'transfer', engraving: true },
      items: [{ product_name: '가위', quantity: 1, unit_price: 100, total_price: 100 }],
    }));
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    const c = ctx.db.rows('contracts')[0];
    expect([c.engraving, c.engraving_status, c.customer_id, c.status]).toEqual([true, 'pending', 'cu9', 'draft']);
  });
});

describe('계약서 → 판매 연결', () => {
  it('계약서로 판매를 만들면 계약서에 판매가 이어지고, 판매를 취소하면 풀린다', async () => {
    seed({ signature_data: SIG, status: 'signed' });
    ctx.db.seed('products', [{ id: 'p1', name: '블런트 6.0', sku: 'B60', stock_quantity: 3, raw_stock: 3, imweb_product_no: null }]);
    const { POST } = await import('@/app/api/sales/route');
    const res = await POST(jsonReq('/api/sales', 'POST', {
      sale: { customer_name: '김점검', sale_date: '2026-11-22', total_amount: 500000, discount_amount: 20000, paid_amount: 480000, payment_method: 'transfer', payment_status: 'paid', contract_id: 'ct1' },
      items: [{ product_id: 'p1', product_name: '블런트 6.0', quantity: 1, unit_price: 500000, total_price: 500000 }],
    }));
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    await settle();
    const sale = ctx.db.rows('offline_sales')[0];
    expect(ctx.db.one('contracts', 'ct1')!.offline_sale_id).toBe(sale.id);

    const { PATCH } = await import('@/app/api/sales/[id]/route');
    await PATCH(jsonReq(`/api/sales/${sale.id}`, 'PATCH', { action: 'cancel', reason: '점검' }), idParam(sale.id));
    await settle();
    expect(ctx.db.one('contracts', 'ct1')!.offline_sale_id).toBeNull();
  });
});

describe('작은 규칙', () => {
  it('이름 가리기 · 문구 · 각인 다듬기 · 서명 검사', () => {
    expect([maskName('김'), maskName('김철'), maskName('김철수'), maskName('남궁민수')]).toEqual(['김', '김*', '김*수', '남**수']);
    expect(signMessage('김점검', T)).toContain(signUrl(T));
    expect(signMessage('', T, { engravingOnly: true })).toContain('각인');
    expect(cleanEngraving('  a \n b  ')).toBe('a b');
    expect(cleanEngraving('x'.repeat(50))).toHaveLength(30);
    expect(isValidSignature(SIG)).toBe(true);
    expect(isValidSignature('data:image/jpeg;base64,' + 'A'.repeat(400))).toBe(false);
  });
});
