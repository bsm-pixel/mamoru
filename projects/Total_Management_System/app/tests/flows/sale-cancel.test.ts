import { describe, it, expect } from 'vitest';
import { ctx, jsonReq, idParam, settle } from '../helpers/setup';
import { saleNet, saleOutstanding } from '@/lib/sales/amounts';

/** ① 판매 등록 → 취소 → 매출 합계 · 고객 미수금 · 재고가 원래대로 */
describe('판매 등록 → 취소', () => {
  it('재고·미수금·매출이 등록 전으로 돌아온다', async () => {
    ctx.db.seed('products', [{ id: 'p1', name: '블런트 5.5', sku: 'B55', stock_quantity: 5, raw_stock: 5, imweb_product_no: null }]);
    ctx.db.seed('customers', [{ id: 'c1', name: '점검고객', phone: '01000000000', outstanding_balance: 0 }]);

    const { POST } = await import('@/app/api/sales/route');
    const res = await POST(jsonReq('/api/sales', 'POST', {
      sale: { customer_id: 'c1', customer_name: '점검고객', customer_phone: '01000000000', sale_date: '2026-11-15',
        total_amount: 300000, discount_amount: 20000, paid_amount: 100000, payment_method: 'cash', payment_status: 'partial' },
      items: [{ product_id: 'p1', product_name: '블런트 5.5', quantity: 2, unit_price: 150000, total_price: 300000 }],
    }));
    const body = await res.json();
    expect(res.status, JSON.stringify(body)).toBe(200);
    await settle();

    // 등록 직후
    const sale = ctx.db.rows('offline_sales')[0];
    expect(sale.sale_number).toBe('OS-20261115-001');
    expect(ctx.db.rows('offline_sale_items')).toHaveLength(1);
    expect(ctx.db.one('products', 'p1')!.stock_quantity).toBe(3);
    expect(ctx.db.one('customers', 'c1')!.outstanding_balance).toBe(180000); // 300000-20000-100000
    expect(saleOutstanding(sale)).toBe(180000);

    // 취소
    const { PATCH } = await import('@/app/api/sales/[id]/route');
    const res2 = await PATCH(jsonReq(`/api/sales/${sale.id}`, 'PATCH', { action: 'cancel', reason: '자동 점검' }), idParam(sale.id));
    const body2 = await res2.json();
    expect(res2.status, JSON.stringify(body2)).toBe(200);
    await settle();

    const after = ctx.db.one('offline_sales', sale.id)!;
    expect(after.cancelled_at).toBeTruthy();
    expect(ctx.db.one('products', 'p1')!.stock_quantity).toBe(5);
    expect(ctx.db.one('products', 'p1')!.raw_stock).toBe(5);
    expect(ctx.db.one('customers', 'c1')!.outstanding_balance).toBe(0);
    // 매출 합계 = 취소 안 된 판매만
    const revenue = ctx.db.rows('offline_sales').filter((s) => !s.cancelled_at).reduce((a, s) => a + saleNet(s), 0);
    expect(revenue).toBe(0);
    // 알림톡은 실제로 나가지 않는다
    expect(ctx.blocked).toEqual([]);
  });
});
