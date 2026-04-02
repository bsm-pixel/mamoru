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

    // 5) COGS / 마진 분석 — 판매 품목별 매입원가 계산
    const saleIds = sales.map((s: { id: string }) => s.id);
    let margin = { total_cogs: 0, gross_profit: 0, margin_rate: 0 };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let byProduct: Array<{ product_id: string; product_name: string; sku: string; qty: number; revenue: number; cogs: number; profit: number; margin_rate: number }> = [];

    if (saleIds.length > 0) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data: saleItems } = await (supabase as any)
        .from('offline_sale_items')
        .select('sale_id, product_id, product_name, sku, quantity, unit_price, total_price')
        .in('sale_id', saleIds);

      // 제품별 매입가 조회
      const productIds = [...new Set((saleItems || []).map((i: { product_id: string }) => i.product_id).filter(Boolean))];
      const purchasePriceMap: Record<string, number> = {};

      if (productIds.length > 0) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const { data: products } = await (supabase as any)
          .from('products')
          .select('id, price_purchase')
          .in('id', productIds);

        for (const p of (products || [])) {
          purchasePriceMap[p.id] = p.price_purchase || 0;
        }
      }

      // 제품별 집계
      const productAgg: Record<string, { product_name: string; sku: string; qty: number; revenue: number; cogs: number }> = {};
      let totalCogs = 0;

      for (const item of (saleItems || [])) {
        const purchasePrice = item.product_id ? (purchasePriceMap[item.product_id] || 0) : 0;
        const itemCogs = purchasePrice * item.quantity;
        totalCogs += itemCogs;

        const key = item.product_id || item.product_name;
        if (!productAgg[key]) {
          productAgg[key] = { product_name: item.product_name, sku: item.sku || '', qty: 0, revenue: 0, cogs: 0 };
        }
        productAgg[key].qty += item.quantity;
        productAgg[key].revenue += item.total_price || (item.unit_price * item.quantity);
        productAgg[key].cogs += itemCogs;
      }

      const grossProfit = saleSummary.total - totalCogs;
      margin = {
        total_cogs: totalCogs,
        gross_profit: grossProfit,
        margin_rate: saleSummary.total > 0 ? Math.round((grossProfit / saleSummary.total) * 1000) / 10 : 0,
      };

      byProduct = Object.entries(productAgg)
        .map(([pid, v]) => ({
          product_id: pid,
          product_name: v.product_name,
          sku: v.sku,
          qty: v.qty,
          revenue: v.revenue,
          cogs: v.cogs,
          profit: v.revenue - v.cogs,
          margin_rate: v.revenue > 0 ? Math.round(((v.revenue - v.cogs) / v.revenue) * 1000) / 10 : 0,
        }))
        .sort((a, b) => b.revenue - a.revenue);
    }

    // 6) 매입처별 지출 집계
    const bySupplier: Record<string, { name: string; total: number; count: number }> = {};
    for (const p of purchases) {
      const name = (p as { supplier_name: string }).supplier_name || '미지정';
      if (!bySupplier[name]) bySupplier[name] = { name, total: 0, count: 0 };
      bySupplier[name].total += (p as { total_amount: number }).total_amount;
      bySupplier[name].count++;
    }
    const supplierRanking = Object.values(bySupplier).sort((a, b) => b.total - a.total);

    // 7) 미지급금 (미완료 PO의 매입처별 잔금)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data: payablesRaw } = await (supabase as any)
      .from('purchase_orders')
      .select('id, supplier_name, total_amount, deposit_amount, balance_amount, status')
      .in('status', ['ordered', 'deposit_paid', 'received']);

    const payablesMap: Record<string, { name: string; total_owed: number; count: number }> = {};
    for (const po of (payablesRaw || [])) {
      const owed = (po.balance_amount ?? 0) > 0 ? po.balance_amount : po.total_amount - (po.deposit_amount || 0);
      if (owed <= 0) continue;
      const name = po.supplier_name || '미지정';
      if (!payablesMap[name]) payablesMap[name] = { name, total_owed: 0, count: 0 };
      payablesMap[name].total_owed += owed;
      payablesMap[name].count++;
    }
    const payables = Object.values(payablesMap).sort((a, b) => b.total_owed - a.total_owed);
    const totalPayables = payables.reduce((s, p) => s + p.total_owed, 0);

    // 8) 미수금 (고객별 outstanding_balance > 0)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data: receivablesRaw } = await (supabase as any)
      .from('customers')
      .select('id, name, company_name, outstanding_balance')
      .gt('outstanding_balance', 0)
      .order('outstanding_balance', { ascending: false });

    // 미결제 판매 건에서 가장 오래된 날짜 조회 (에이징용)
    const receivableIds = (receivablesRaw || []).map((c: { id: string }) => c.id);
    let oldestSaleMap: Record<string, string> = {};
    if (receivableIds.length > 0) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data: unpaidSales } = await (supabase as any)
        .from('offline_sales')
        .select('customer_id, sale_date')
        .in('customer_id', receivableIds)
        .in('payment_status', ['unpaid', 'partial'])
        .is('cancelled_at', null)
        .order('sale_date', { ascending: true });
      for (const s of (unpaidSales || [])) {
        if (!oldestSaleMap[s.customer_id]) oldestSaleMap[s.customer_id] = s.sale_date;
      }
    }

    const now = new Date();
    const receivables = (receivablesRaw || []).map((c: { id: string; name: string; company_name: string; outstanding_balance: number }) => {
      const oldest = oldestSaleMap[c.id];
      const daysOverdue = oldest ? Math.floor((now.getTime() - new Date(oldest).getTime()) / (1000 * 60 * 60 * 24)) : 0;
      const aging = daysOverdue <= 30 ? '30일이내' : daysOverdue <= 60 ? '30~60일' : daysOverdue <= 90 ? '60~90일' : '90일초과';
      return { id: c.id, name: c.company_name || c.name, outstanding: c.outstanding_balance, daysOverdue, aging };
    });
    const totalReceivables = receivables.reduce((s: number, r: { outstanding: number }) => s + r.outstanding, 0);

    // 에이징 요약
    type RItem = { outstanding: number; daysOverdue: number };
    const agingSummary = {
      within30: receivables.filter((r: RItem) => r.daysOverdue <= 30).reduce((s: number, r: RItem) => s + r.outstanding, 0),
      d30to60: receivables.filter((r: RItem) => r.daysOverdue > 30 && r.daysOverdue <= 60).reduce((s: number, r: RItem) => s + r.outstanding, 0),
      d60to90: receivables.filter((r: RItem) => r.daysOverdue > 60 && r.daysOverdue <= 90).reduce((s: number, r: RItem) => s + r.outstanding, 0),
      over90: receivables.filter((r: RItem) => r.daysOverdue > 90).reduce((s: number, r: RItem) => s + r.outstanding, 0),
    };

    // 9) 복원수리 매출 집계 (paid_at 기준)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data: repairSalesRaw } = await (supabase as any)
      .from('repairs')
      .select('id, as_id, name, phone, service_cost, shipping_fee, total_amount, paid_at, created_at')
      .not('paid_at', 'is', null)
      .gte('paid_at', `${fromDate}T00:00:00`)
      .lte('paid_at', `${toDate}T23:59:59`)
      .not('status', 'eq', 'cancelled');

    const repairSales = repairSalesRaw || [];
    const repairSalesSummary = {
      count: repairSales.length,
      total: repairSales.reduce((s: number, r: { total_amount: number }) => s + (r.total_amount || 0), 0),
      service_cost_total: repairSales.reduce((s: number, r: { service_cost: number }) => s + (r.service_cost || 0), 0),
      shipping_fee_total: repairSales.reduce((s: number, r: { shipping_fee: number }) => s + (r.shipping_fee || 0), 0),
    };

    // 복원수리 일별 매출
    const dailyRepairs: Record<string, number> = {};
    for (const r of repairSales) {
      const d = (r.paid_at as string).slice(0, 10);
      dailyRepairs[d] = (dailyRepairs[d] || 0) + (r.total_amount || 0);
    }

    const totalRevenue = saleSummary.total + repairSalesSummary.total;

    return NextResponse.json({
      period: { from: fromDate, to: toDate },
      sales: saleSummary,
      repair_sales: repairSalesSummary,
      total_revenue: totalRevenue,
      purchases: purchaseSummary,
      vat: vatSummary,
      margin,
      by_product: byProduct,
      by_supplier: supplierRanking,
      payables: { items: payables, total: totalPayables },
      receivables: { items: receivables, total: totalReceivables, aging: agingSummary },
      daily: { sales: dailySales, purchases: dailyPurchases, repairs: dailyRepairs },
      details: { sales, purchases, repairs: repairSales },
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
