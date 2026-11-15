import { describe, it, expect } from 'vitest';
import { phoneError, emailError, amountError, typingPhone } from '@/lib/utils/validate';
import { ctx, jsonReq, idParam } from '../helpers/setup';

/** 4차 4·5번 — 입력 검사 (화면·서버 공용) */
describe('입력 검사 규칙', () => {
  it('전화번호', () => {
    for (const ok of ['', '010-1234-5678', '01012345678', '02-123-4567', '031-123-4567', '0507-1234-5678', '+86 138 0013 8000']) expect(phoneError(ok), ok).toBeNull();
    for (const bad of ['abc123', '1234', '010-12', '12345678901234', '910-1234-5678']) expect(phoneError(bad), bad).toBeTruthy();
  });
  it('이메일', () => {
    expect(emailError('')).toBeNull();
    expect(emailError('a@b.kr')).toBeNull();
    for (const bad of ['notanemail', 'a@b', '@b.com', 'a b@c.com']) expect(emailError(bad), bad).toBeTruthy();
  });
  it('금액', () => {
    expect(amountError(0, '소매가', { required: true })).toBeNull();
    expect(amountError('', '소매가', { required: true })).toBeTruthy();
    expect(amountError(-100, '소매가')).toBeTruthy();
    expect(amountError('abc', '소매가')).toBeTruthy();
    expect(amountError(undefined, '딜러가')).toBeNull();
  });
  it('자동 하이픈', () => {
    expect(typingPhone('01012345678')).toBe('010-1234-5678');
    expect(typingPhone('0212345678')).toBe('02-1234-5678');
    expect(typingPhone('021234567')).toBe('02-123-4567');
    expect(typingPhone('0311234567')).toBe('031-123-4567');
    expect(typingPhone('050712345678')).toBe('0507-1234-5678');
    expect(typingPhone('+8613800138000')).toBe('+8613800138000');
  });
});

describe('서버가 잘못된 입력을 저장하지 않는다', () => {
  it('고객: 이상한 연락처·이메일은 400, 저장 0건', async () => {
    const { POST } = await import('@/app/api/customers/route');
    const a = await POST(jsonReq('/api/customers', 'POST', { name: '점검', phone: 'abc123' }));
    const b = await POST(jsonReq('/api/customers', 'POST', { name: '점검', phone: '010-1111-2222', email: 'notanemail' }));
    expect([a.status, b.status]).toEqual([400, 400]);
    expect(ctx.db.rows('customers')).toHaveLength(0);
  });
  it('고객 수정: 연락처를 이상하게 바꾸면 400, 원래 값 유지', async () => {
    ctx.db.seed('customers', [{ id: 'c1', name: '점검', phone: '010-1111-2222' }]);
    const { PATCH } = await import('@/app/api/customers/[id]/route');
    const res = await PATCH(jsonReq('/api/customers/c1', 'PATCH', { phone: 'abc123' }), idParam('c1'));
    expect(res.status).toBe(400);
    expect(ctx.db.one('customers', 'c1')!.phone).toBe('010-1111-2222');
  });
  it('제품: 음수 가격·빈 소매가는 400, 저장 0건 / 정상 값은 저장', async () => {
    const { POST } = await import('@/app/api/products/route');
    const neg = await POST(jsonReq('/api/products', 'POST', { sku: 'T1', name: '점검', price: -100 }));
    const empty = await POST(jsonReq('/api/products', 'POST', { sku: 'T2', name: '점검', price: '' }));
    expect([neg.status, empty.status]).toEqual([400, 400]);
    expect(ctx.db.rows('products')).toHaveLength(0);
    const ok = await POST(jsonReq('/api/products', 'POST', { sku: 'T3', name: '점검', price: 150000 }));
    expect(ok.status, JSON.stringify(await ok.clone().json())).toBe(200);
    expect(ctx.db.rows('products')).toHaveLength(1);
  });
});
