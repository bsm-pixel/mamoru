import { describe, it, expect } from 'vitest';
import { ctx, jsonReq, idParam } from '../helpers/setup';

/** 4차 6번 — 고객 상세 「판매 N건 · 매출」은 취소·반품 건을 세지 않는다 */
describe('고객 상세 합계', () => {
  it('취소·반품한 판매는 건수와 매출에서 빠지고, 21건이 넘어도 전부 센다', async () => {
    ctx.db.seed('customers', [{ id: 'c1', name: '점검고객', phone: '01000000000' }]);
    const base = { customer_id: 'c1', payment_method: 'cash', payment_status: 'paid', total_amount: 10000, cancelled_at: null as string | null, returned_at: null as string | null };
    const rows = Array.from({ length: 23 }, (_, i) => ({
      ...base, id: `s${i}`, sale_number: `OS-${i}`, sale_date: `2026-10-${String(i + 1).padStart(2, '0')}`, paid_amount: 10000,
    }));
    rows.push({ ...base, id: 'sc', sale_number: 'OS-C', sale_date: '2026-11-10', paid_amount: 999, cancelled_at: '2026-11-11T00:00:00Z' });
    rows.push({ ...base, id: 'sr', sale_number: 'OS-R', sale_date: '2026-11-12', paid_amount: 777, returned_at: '2026-11-13T00:00:00Z' });
    ctx.db.seed('offline_sales', rows);

    const { GET } = await import('@/app/api/customers/[id]/route');
    const res = await GET(jsonReq('/api/customers/c1', 'GET'), idParam('c1'));
    const body = await res.json();
    expect(res.status, JSON.stringify(body).slice(0, 300)).toBe(200);
    expect(body.summary.totalSales).toBe(23);
    expect(body.summary.totalSalesAmount).toBe(230000);
    expect(body.summary.lastSaleDate).toBe('2026-10-23'); // 취소·반품 건의 날짜는 '마지막 거래'가 아니다
    expect(body.sales.length).toBeLessThanOrEqual(20);     // 이력 목록은 최근 20건 그대로
    expect(body.sales.find((s: { id: string }) => s.id === 'sc').cancelled_at).toBeTruthy(); // 화면이 '취소됨'으로 그릴 수 있게
  });
});
