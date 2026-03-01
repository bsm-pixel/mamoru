import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

/**
 * GET /api/reports/summary — 매출/매입/VAT 기간별 집계
 * ?from=2026-01-01&to=2026-01-31
 */
export async function GET(req: NextRequest) {
  const supabase = createServiceClient();
  const sp = req.nextUrl.searchParams;
  const fromDate = sp.get('from') || new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString().slice(0, 10);
  const toDate = sp.get('to') || new Date().toISOString().slice(0, 10);

  try {
    // 1) 매출 집계 (offline_sales)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data: salesRaw } = await (supabase as any)
      .from('offline_sales')
      .select('id, sale_date, total_amount, supply_amount, vat_amount, payment_method, payment_status, customer_name, customer_id, discount_amount, paid_amount')
      .gte('sale_date', fromDate)
      .lte('sale_date', toDate)
      .order('sale_date', { ascending: false });

    const sales = salesRaw || [];

    const saleSummary = {
      count: sales.length,
      total: sales.reduce((s: number, r: { total_amount: number }) => s + r.total_amount, 0),
      supply: sales.reduce((s: number, r: { supply_amount: number }) => s + (r.supply_amount || 0), 0),
      vat: sales.reduce((s: number, r: { vat_amount: number }) => s + (r.vat_amount || 0), 0),
      discount: sales.reduce((s: number, r: { discount_amount: number }) => s + (r.discount_amount || 0), 0),
      paid: sales.reduce((s: number, r: { paid_amount: number }) => s + (r.paid_amount || 0), 0),
      by_method: groupBy(sales, 'payment_method', 'total_amount'),
    };

    // 2) 매입 집계 (purchase_orders - received/balance_paid만)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data: poRaw } = await (supabase as any)
      .from('purchase_orders')
      .select('id, order_date, total_amount, supply_amount, vat_amount, deposit_amount, balance_amount, status, supplier_name')
      .gte('order_date', fromDate)
      .lte('order_date', toDate)
      .in('status', ['received', 'balance_paid', 'deposit_paid', 'ordered'])
      .order('order_date', { ascending: false });

    const purchases = poRaw || [];

    const purchaseSummary = {
      count: purchases.length,
      total: purchases.reduce((s: number, r: { total_amount: number }) => s + r.total_amount, 0),
      supply: purchases.reduce((s: number, r: { supply_amount: number }) => s + (r.supply_amount || 0), 0),
      vat: purchases.reduce((s: number, r: { vat_amount: number }) => s + (r.vat_amount || 0), 0),
      deposit: purchases.reduce((s: number, r: { deposit_amount: number }) => s + (r.deposit_amount || 0), 0),
    };

    // 3) 일별 매출 추이 (차트용)
    const dailySales: Record<string, number> = {};
    const dailyPurchases: Record<string, number> = {};

    for (const s of sales) {
      const d = (s as { sale_date: string }).sale_date;
      dailySales[d] = (dailySales[d] || 0) + (s as { total_amount: number }).total_amount;
    }
    for (const p of purchases) {
      const d = (p as { order_date: string }).order_date;
      dailyPurchases[d] = (dailyPurchases[d] || 0) + (p as { total_amount: number }).total_amount;
    }

    // 4) VAT 요약 (매출세액 - 매입세액)
    const vatSummary = {
      sales_vat: saleSummary.vat,
      purchase_vat: purchaseSummary.vat,
      net_vat: saleSummary.vat - purchaseSummary.vat,
    };

    return NextResponse.json({
      period: { from: fromDate, to: toDate },
      sales: saleSummary,
      purchases: purchaseSummary,
      vat: vatSummary,
      daily: { sales: dailySales, purchases: dailyPurchases },
      details: { sales, purchases },
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

function groupBy(arr: Record<string, unknown>[], key: string, sumKey: string): Record<string, number> {
  const result: Record<string, number> = {};
  for (const item of arr) {
    const k = (item[key] as string) || 'unknown';
    result[k] = (result[k] || 0) + ((item[sumKey] as number) || 0);
  }
  return result;
}
