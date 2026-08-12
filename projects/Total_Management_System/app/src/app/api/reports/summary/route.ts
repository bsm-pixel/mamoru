import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

/**
 * GET /api/reports/summary — 매출/매입/VAT 기간별 집계
 * ?from=2026-01-01&to=2026-01-31
 *
 * 2단계 매출 3분할 (2026-05-12):
 *   총매출 = 제품 매출(B2C + B2B, 납품 포함, RS 제외) + 복원수리 매출(A 접수 + B 판매RS + C 납품RS)
 *   - 제품 매출: offline_sales (total−discount) − 그 주문 RS items + deliveries (total−discount) − 납품 RS items
 *   - 복원수리: repairs(접수시스템, paid_at 기준 실제 금액) + offline_sale_items(category='RS') + delivery_items(category='RS')
 *   - by_product / margin: RS 항목 제외 (RS 는 제품이 아니므로 제품 랭킹·원가 집계에서 빠짐)
 */
export async function GET(req: NextRequest) {
  const supabase = createServiceClient();
  const sp = req.nextUrl.searchParams;
  const fromDate = sp.get('from') || new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString().slice(0, 10);
  const toDate = sp.get('to') || new Date().toISOString().slice(0, 10);

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const isB2BCt = (ct: string | null | undefined) => ct === 'dealer' || ct === 'academy';

    // ─── 1) 오프라인 판매 (offline_sales + items) — 취소 제외 / 반품은 '판매월'에 인식(반품시점 반영, 2026-08-01) ───
    //   반품(returned)은 여기서 제외하지 않는다 → 판매월 매출로 잡히고, 반품월엔 아래 retSales 로 −차감.
    const { data: salesRaw } = await db
      .from('offline_sales')
      .select('id, sale_date, total_amount, supply_amount, vat_amount, payment_method, payment_status, customer_name, customer_id, customer_type, discount_amount, paid_amount, payment_detail')
      .gte('sale_date', fromDate)
      .lte('sale_date', toDate)
      .is('cancelled_at', null)
      .order('sale_date', { ascending: false });
    const sales: Array<{ id: string; sale_date: string; total_amount: number; supply_amount: number; vat_amount: number; payment_method: string; payment_status: string; customer_name: string; customer_id: string | null; customer_type: string | null; discount_amount: number; paid_amount: number; payment_detail: Record<string, number> | null }> = salesRaw || [];
    const saleIds = sales.map((s) => s.id);

    let saleItems: Array<{ sale_id: string; product_id: string | null; product_name: string; sku: string | null; quantity: number; unit_price: number; total_price: number; category: string | null }> = [];
    if (saleIds.length > 0) {
      const { data } = await db
        .from('offline_sale_items')
        .select('sale_id, product_id, product_name, sku, quantity, unit_price, total_price, category')
        .in('sale_id', saleIds);
      saleItems = data || [];
    }
    // 오프라인 복원수리(RS) 항목 — 0원 무상 제외, sale_id 별 합계 (배송비는 금액엔 포함, 자루 수엔 제외)
    const offlineRsBySale: Record<string, { amount: number; qty: number }> = {};
    for (const it of saleItems) {
      if (it.category === 'RS' && (it.total_price || 0) > 0) {
        if (!offlineRsBySale[it.sale_id]) offlineRsBySale[it.sale_id] = { amount: 0, qty: 0 };
        offlineRsBySale[it.sale_id].amount += it.total_price || 0;
        if (it.product_name !== '배송비') offlineRsBySale[it.sale_id].qty += it.quantity || 0;
      }
    }
    const productSaleItems = saleItems.filter((it) => it.category !== 'RS'); // 제품 항목만 (COGS·랭킹용)

    // ─── 반품(반품시점 반영) — returned_at 이 이 기간인 판매를 −차감 ───
    //   판매 자체는 위 sales(판매월)에 이미 인식됨. 여기선 반품월에 음수로 빼서 "반품 시점 반영".
    //   판매·반품이 같은 기간이면 +와 −가 상쇄(net 0). 다른 기간이면 각 월에 정확히 귀속.
    const { data: retRaw } = await db
      .from('offline_sales')
      .select('id, returned_at, total_amount, discount_amount, customer_type')
      .gte('returned_at', `${fromDate}T00:00:00`)
      .lte('returned_at', `${toDate}T23:59:59`)
      .is('cancelled_at', null);
    const retSales: Array<{ id: string; returned_at: string; total_amount: number; discount_amount: number; customer_type: string | null }> = retRaw || [];
    let retItems: typeof saleItems = [];
    if (retSales.length > 0) {
      const { data } = await db
        .from('offline_sale_items')
        .select('sale_id, product_id, product_name, sku, quantity, unit_price, total_price, category')
        .in('sale_id', retSales.map((s) => s.id));
      retItems = data || [];
    }
    const retRsBySale: Record<string, { amount: number; qty: number }> = {};
    for (const it of retItems) {
      if (it.category === 'RS' && (it.total_price || 0) > 0) {
        if (!retRsBySale[it.sale_id]) retRsBySale[it.sale_id] = { amount: 0, qty: 0 };
        retRsBySale[it.sale_id].amount += it.total_price || 0;
        if (it.product_name !== '배송비') retRsBySale[it.sale_id].qty += it.quantity || 0;
      }
    }
    const retProductItems = retItems.filter((it) => it.category !== 'RS'); // 반품 제품 항목(−) — allProductItems 에 음수로
    const retSaleB2BMap: Record<string, boolean> = {};
    for (const s of retSales) retSaleB2BMap[s.id] = isB2BCt(s.customer_type);

    // 반품 집계 (반품월 −)
    let retProductB2C = 0, retProductB2BOffline = 0, retSaleTotal = 0;
    const retDailyProduct: Record<string, number> = {};
    const retDailyRepairs: Record<string, number> = {};
    for (const s of retSales) {
      const rs = retRsBySale[s.id]?.amount || 0;
      const prod = (s.total_amount || 0) - (s.discount_amount || 0) - rs;
      if (isB2BCt(s.customer_type)) retProductB2BOffline += prod; else retProductB2C += prod;
      retSaleTotal += s.total_amount || 0;
      const d = (s.returned_at || '').slice(0, 10);
      if (prod !== 0) retDailyProduct[d] = (retDailyProduct[d] || 0) + prod;
      if (rs > 0) retDailyRepairs[d] = (retDailyRepairs[d] || 0) + rs;
    }
    const retRsTotal = Object.values(retRsBySale).reduce((s, v) => s + v.amount, 0);
    const retRsQty = Object.values(retRsBySale).reduce((s, v) => s + v.qty, 0);

    // 순(net) — 판매월 − 반품월
    const offlineRsTotal = Object.values(offlineRsBySale).reduce((s, v) => s + v.amount, 0) - retRsTotal;
    const offlineRsQty = Object.values(offlineRsBySale).reduce((s, v) => s + v.qty, 0) - retRsQty;
    const offlineRsCount = Object.keys(offlineRsBySale).length; // 건수는 판매월 기준(호환용)

    // 제품 매출 B2C/B2B 분리 (판매월 +) 이후 (반품월 −) 차감
    let productB2C = 0, productB2BOffline = 0;
    for (const s of sales) {
      const rs = offlineRsBySale[s.id]?.amount || 0;
      const prod = (s.total_amount || 0) - (s.discount_amount || 0) - rs;
      if (isB2BCt(s.customer_type)) productB2BOffline += prod; else productB2C += prod;
    }
    productB2C -= retProductB2C;
    productB2BOffline -= retProductB2BOffline;

    // ─── 2) 납품 (deliveries + items) — 전부 B2B 거래처 취급 ───
    const { data: dlRaw } = await db
      .from('deliveries')
      .select('id, delivery_date, total_amount, discount_amount, customer_name, status')
      .gte('delivery_date', fromDate)
      .lte('delivery_date', toDate)
      .in('status', ['confirmed', 'shipped', 'settled'])
      .is('cancelled_at', null)
      .order('delivery_date', { ascending: false });
    const deliveries: Array<{ id: string; delivery_date: string; total_amount: number; discount_amount: number; customer_name: string; status: string }> = dlRaw || [];
    const dlIds = deliveries.map((d) => d.id);
    let dlItems: Array<{ delivery_id: string; product_id: string | null; product_name: string; sku: string | null; quantity: number; unit_price: number; total_price: number; category: string | null }> = [];
    if (dlIds.length > 0) {
      const { data } = await db
        .from('delivery_items')
        .select('delivery_id, product_id, product_name, sku, quantity, unit_price, total_price, category')
        .in('delivery_id', dlIds);
      dlItems = data || [];
    }
    const dlRsByDelivery: Record<string, { amount: number; qty: number }> = {};
    for (const it of dlItems) {
      if (it.category === 'RS' && (it.total_price || 0) > 0) {
        if (!dlRsByDelivery[it.delivery_id]) dlRsByDelivery[it.delivery_id] = { amount: 0, qty: 0 };
        dlRsByDelivery[it.delivery_id].amount += it.total_price || 0;
        if (it.product_name !== '배송비') dlRsByDelivery[it.delivery_id].qty += it.quantity || 0;
      }
    }
    const productDlItems = dlItems.filter((it) => it.category !== 'RS');
    const deliveryRsTotal = Object.values(dlRsByDelivery).reduce((s, v) => s + v.amount, 0);
    const deliveryRsQty = Object.values(dlRsByDelivery).reduce((s, v) => s + v.qty, 0);
    const deliveryRsCount = Object.keys(dlRsByDelivery).length;
    let productB2BDelivery = 0;
    for (const d of deliveries) {
      const rs = dlRsByDelivery[d.id]?.amount || 0;
      productB2BDelivery += (d.total_amount || 0) - rs; // 납품 total은 이미 net(할인 재차감 금지)
    }
    const productB2B = productB2BOffline + productB2BDelivery;
    const productTotal = productB2C + productB2B;

    // 오프라인 판매 전체 요약 (RS 포함 — 호환용 "오프라인 판매 총액")
    const saleSummary = {
      count: sales.length,
      total: sales.reduce((s, r) => s + (r.total_amount || 0), 0) - retSaleTotal,  // 반품 −반영(net)
      supply: sales.reduce((s, r) => s + (r.supply_amount || 0), 0),
      vat: sales.reduce((s, r) => s + (r.vat_amount || 0), 0),
      discount: sales.reduce((s, r) => s + (r.discount_amount || 0), 0),
      paid: sales.reduce((s, r) => s + (r.paid_amount || 0), 0),
      by_method: buildByMethod(sales),
    };

    // ─── 3) 매입 집계 (purchase_orders) ───
    const { data: poRaw } = await db
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

    // ─── 4) 일별 매출/매입 추이 (차트용) ───
    // dailySales(호환): offline_sales total / dailyProduct: 제품(판매−RS + 납품제품) — 탭별 차트용 / 매입: order_date
    const dailySales: Record<string, number> = {};
    const dailyProduct: Record<string, number> = {};
    const dailyPurchases: Record<string, number> = {};
    for (const s of sales) {
      dailySales[s.sale_date] = (dailySales[s.sale_date] || 0) + (s.total_amount || 0);
      const rs = offlineRsBySale[s.id]?.amount || 0;
      const prod = (s.total_amount || 0) - (s.discount_amount || 0) - rs;
      if (prod !== 0) dailyProduct[s.sale_date] = (dailyProduct[s.sale_date] || 0) + prod;
    }
    for (const d of deliveries) {
      const rs = dlRsByDelivery[d.id]?.amount || 0;
      const prod = (d.total_amount || 0) - rs; // 납품 total은 이미 net
      if (prod !== 0) dailyProduct[d.delivery_date] = (dailyProduct[d.delivery_date] || 0) + prod;
    }
    // 반품 제품 매출(−) 을 반품월(returned_at)에 차감 (반품시점 반영)
    for (const d of Object.keys(retDailyProduct)) dailyProduct[d] = (dailyProduct[d] || 0) - retDailyProduct[d];
    for (const p of purchases) {
      const d = (p as { order_date: string }).order_date;
      dailyPurchases[d] = (dailyPurchases[d] || 0) + (p as { total_amount: number }).total_amount;
    }

    // ─── 5) VAT 요약 (매출세액 − 매입세액) ───
    const vatSummary = {
      sales_vat: saleSummary.vat,
      purchase_vat: purchaseSummary.vat,
      net_vat: saleSummary.vat - purchaseSummary.vat,
    };

    // ─── 5-2) 아임웹 온라인 주문 항목 (품목별 매출/랭킹 포함용) — 취소/환불 제외, 배송비·0원 제외 ───
    const { data: ordRaw } = await db
      .from('orders')
      .select('id, ordered_at, status')
      .gte('ordered_at', `${fromDate}T00:00:00`)
      .lte('ordered_at', `${toDate}T23:59:59`)
      .not('status', 'in', '("cancelled","refunded")');
    const orderIds = (ordRaw || []).map((o: { id: string }) => o.id);
    let productOrderItems: Array<{ product_id: string | null; product_name: string; sku: string | null; quantity: number; unit_price: number; total_price: number }> = [];
    if (orderIds.length > 0) {
      const { data: oiRaw } = await db
        .from('order_items')
        .select('order_id, product_id, product_name, quantity, unit_price, total_price')
        .in('order_id', orderIds);
      productOrderItems = (oiRaw || [])
        .filter((it: { product_name: string; total_price: number; unit_price: number; quantity: number }) =>
          it.product_name !== '배송비' && ((it.total_price || (it.unit_price || 0) * (it.quantity || 0)) > 0))
        .map((it: { product_id: string | null; product_name: string; quantity: number; unit_price: number; total_price: number }) =>
          ({ product_id: it.product_id, product_name: it.product_name, sku: null, quantity: it.quantity, unit_price: it.unit_price, total_price: it.total_price }));
    }

    // 오프라인 판매 B2B 여부 (customer_type dealer/academy) — 품목 B2C/B2B 분리용
    const saleB2BMap: Record<string, boolean> = {};
    for (const s of sales) saleB2BMap[s.id] = isB2BCt(s.customer_type);

    // ─── 6) COGS / 마진 / 제품 랭킹 — 제품 항목만 (RS 제외), 오프라인 + 납품 + 온라인 합산 ───
    //   channel: 오프라인=고객유형 / 납품=b2b / 온라인=b2c
    const allProductItems: Array<{ product_id: string | null; product_name: string; sku: string | null; quantity: number; unit_price: number; total_price: number; channel: 'b2c' | 'b2b' }> = [
      ...productSaleItems.map((it) => ({ product_id: it.product_id, product_name: it.product_name, sku: it.sku, quantity: it.quantity, unit_price: it.unit_price, total_price: it.total_price, channel: (saleB2BMap[it.sale_id] ? 'b2b' : 'b2c') as 'b2c' | 'b2b' })),
      // 반품 제품 항목(−) — 수량·매출·원가를 음수로 넣어 by_product/마진이 자동으로 net 됨(반품시점 반영)
      ...retProductItems.map((it) => ({ product_id: it.product_id, product_name: it.product_name, sku: it.sku, quantity: -(it.quantity || 0), unit_price: it.unit_price, total_price: -(it.total_price || 0), channel: (retSaleB2BMap[it.sale_id] ? 'b2b' : 'b2c') as 'b2c' | 'b2b' })),
      ...productDlItems.map((it) => ({ product_id: it.product_id, product_name: it.product_name, sku: it.sku, quantity: it.quantity, unit_price: it.unit_price, total_price: it.total_price, channel: 'b2b' as const })),
      ...productOrderItems.map((it) => ({ ...it, channel: 'b2c' as const })),
    ];
    const productIds = [...new Set(allProductItems.map((i) => i.product_id).filter(Boolean))] as string[];
    const purchasePriceMap: Record<string, number> = {};
    const skuMap: Record<string, string> = {};
    const categoryMap: Record<string, string> = {};
    if (productIds.length > 0) {
      const { data: products } = await db.from('products').select('id, price_purchase, sku, category').in('id', productIds);
      for (const p of (products || [])) { purchasePriceMap[p.id] = p.price_purchase || 0; if (p.sku) skuMap[p.id] = p.sku; if (p.category) categoryMap[p.id] = p.category; }
    }
    const productAgg: Record<string, { product_name: string; sku: string; qty: number; revenue: number; cogs: number; b2cQty: number; b2cRev: number; b2cCogs: number; b2bQty: number; b2bRev: number; b2bCogs: number }> = {};
    let totalCogs = 0;
    for (const item of allProductItems) {
      const purchasePrice = item.product_id ? (purchasePriceMap[item.product_id] || 0) : 0;
      const itemCogs = purchasePrice * (item.quantity || 0);
      totalCogs += itemCogs;
      const rev = item.total_price || ((item.unit_price || 0) * (item.quantity || 0));
      const key = item.product_id || item.product_name;
      if (!productAgg[key]) productAgg[key] = { product_name: item.product_name, sku: item.sku || '', qty: 0, revenue: 0, cogs: 0, b2cQty: 0, b2cRev: 0, b2cCogs: 0, b2bQty: 0, b2bRev: 0, b2bCogs: 0 };
      const a = productAgg[key];
      a.qty += item.quantity || 0; a.revenue += rev; a.cogs += itemCogs;
      if (item.channel === 'b2b') { a.b2bQty += item.quantity || 0; a.b2bRev += rev; a.b2bCogs += itemCogs; }
      else { a.b2cQty += item.quantity || 0; a.b2cRev += rev; a.b2cCogs += itemCogs; }
    }
    const productGrossProfit = productTotal - totalCogs;
    const margin = {
      total_cogs: totalCogs,
      gross_profit: productGrossProfit,
      margin_rate: productTotal > 0 ? Math.round((productGrossProfit / productTotal) * 1000) / 10 : 0,
    };
    const byProduct = Object.entries(productAgg)
      .map(([pid, v]) => ({
        product_id: pid,
        product_name: v.product_name,
        sku: skuMap[pid] || v.sku,
        category: categoryMap[pid] || null,
        qty: v.qty,
        revenue: v.revenue,
        cogs: v.cogs,
        profit: v.revenue - v.cogs,
        margin_rate: v.revenue > 0 ? Math.round(((v.revenue - v.cogs) / v.revenue) * 1000) / 10 : 0,
        b2c_qty: v.b2cQty, b2c_revenue: v.b2cRev, b2c_cogs: v.b2cCogs,
        b2b_qty: v.b2bQty, b2b_revenue: v.b2bRev, b2b_cogs: v.b2bCogs,
      }))
      .sort((a, b) => b.revenue - a.revenue);

    // ─── 7) 매입처별 지출 ───
    const bySupplier: Record<string, { name: string; total: number; count: number }> = {};
    for (const p of purchases) {
      const name = (p as { supplier_name: string }).supplier_name || '미지정';
      if (!bySupplier[name]) bySupplier[name] = { name, total: 0, count: 0 };
      bySupplier[name].total += (p as { total_amount: number }).total_amount;
      bySupplier[name].count++;
    }
    const supplierRanking = Object.values(bySupplier).sort((a, b) => b.total - a.total);

    // ─── 8) 미지급금 (미완료 PO의 매입처별 잔금) ───
    const { data: payablesRaw } = await db
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

    // ─── 9) 미수금 (고객별 outstanding_balance > 0) + 에이징 ───
    const { data: receivablesRaw } = await db
      .from('customers')
      .select('id, name, company_name, outstanding_balance')
      .gt('outstanding_balance', 0)
      .order('outstanding_balance', { ascending: false });
    const receivableIds = (receivablesRaw || []).map((c: { id: string }) => c.id);
    const oldestSaleMap: Record<string, string> = {};
    if (receivableIds.length > 0) {
      const { data: unpaidSales } = await db
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
    type RItem = { outstanding: number; daysOverdue: number };
    const agingSummary = {
      within30: receivables.filter((r: RItem) => r.daysOverdue <= 30).reduce((s: number, r: RItem) => s + r.outstanding, 0),
      d30to60: receivables.filter((r: RItem) => r.daysOverdue > 30 && r.daysOverdue <= 60).reduce((s: number, r: RItem) => s + r.outstanding, 0),
      d60to90: receivables.filter((r: RItem) => r.daysOverdue > 60 && r.daysOverdue <= 90).reduce((s: number, r: RItem) => s + r.outstanding, 0),
      over90: receivables.filter((r: RItem) => r.daysOverdue > 90).reduce((s: number, r: RItem) => s + r.outstanding, 0),
    };

    // ─── 10) 복원수리 매출 = A(접수 repairs, paid_at 기준 실제 금액) + B(offline RS) + C(delivery RS) ───
    const { data: repairSalesRaw } = await db
      .from('repairs')
      .select('id, as_id, name, phone, service_cost, shipping_fee, total_amount, paid_at, payment_method, created_at')
      .not('paid_at', 'is', null)
      .gte('paid_at', `${fromDate}T00:00:00`)
      .lte('paid_at', `${toDate}T23:59:59`)
      .not('status', 'eq', 'cancelled');
    const repairSales = repairSalesRaw || [];
    const repairAIntake = repairSales.reduce((s: number, r: { total_amount: number }) => s + (r.total_amount || 0), 0);
    const repairServiceCost = repairSales.reduce((s: number, r: { service_cost: number }) => s + (r.service_cost || 0), 0);
    const repairShippingFee = repairSales.reduce((s: number, r: { shipping_fee: number }) => s + (r.shipping_fee || 0), 0);
    const repairTotal = repairAIntake + offlineRsTotal + deliveryRsTotal;
    const repairSalesSummary = {
      total: repairTotal,                       // A + B + C
      a_intake: repairAIntake,                  // 접수시스템(repairs, 입금 기준 실제 금액)
      b_offline_rs: offlineRsTotal,             // 판매시스템 RS (offline_sale_items category='RS')
      c_delivery_rs: deliveryRsTotal,           // 납품 RS (delivery_items category='RS')
      a_count: repairSales.length,
      b_count: offlineRsCount,
      c_count: deliveryRsCount,
      bc_qty: offlineRsQty + deliveryRsQty,     // B+C 자루 수 (A 는 접수 건수만, 자루 정보는 repairs.qty_* — 여기선 미집계)
      service_cost_total: repairServiceCost,    // (A 기준)
      shipping_fee_total: repairShippingFee,    // (A 기준)
      count: repairSales.length + offlineRsCount + deliveryRsCount, // 호환: 전체 건수
      // 120: 결제수단별 집계 (A채널 = 접수 repairs, payment_method 기록된 입금건). 카드 복원수리 회계 반영
      by_method: groupBy(
        repairSales.filter((r: { payment_method?: string | null; total_amount?: number }) => r.payment_method && (r.total_amount || 0) > 0) as unknown as Record<string, unknown>[],
        'payment_method', 'total_amount'
      ),
    };

    // 복원수리 일별 (A: paid_at / B: sale_date / C: delivery_date)
    const dailyRepairs: Record<string, number> = {};
    for (const r of repairSales) {
      const d = (r.paid_at as string).slice(0, 10);
      dailyRepairs[d] = (dailyRepairs[d] || 0) + (r.total_amount || 0);
    }
    for (const s of sales) {
      const rs = offlineRsBySale[s.id]?.amount || 0;
      if (rs > 0) dailyRepairs[s.sale_date] = (dailyRepairs[s.sale_date] || 0) + rs;
    }
    for (const d of deliveries) {
      const rs = dlRsByDelivery[d.id]?.amount || 0;
      if (rs > 0) dailyRepairs[d.delivery_date] = (dailyRepairs[d.delivery_date] || 0) + rs;
    }
    // 반품 RS(−) 을 반품월에 차감
    for (const d of Object.keys(retDailyRepairs)) dailyRepairs[d] = (dailyRepairs[d] || 0) - retDailyRepairs[d];

    // ─── 11) 제품 매출 객체 (RS 제외, 납품 포함, B2C/B2B) ───
    const productSalesSummary = {
      total: productTotal,
      b2c: productB2C,
      b2b: productB2B,
      b2b_offline: productB2BOffline,
      b2b_delivery: productB2BDelivery,
      offline_count: sales.length,
      delivery_count: deliveries.length,
    };

    const totalRevenue = productTotal + repairTotal;

    // ─── 12) 경비 / 손익 ───
    const { data: expensesRaw } = await db
      .from('expenses')
      .select('amount, category')
      .gte('expense_date', fromDate)
      .lte('expense_date', toDate);
    const totalExpenses = (expensesRaw || []).reduce((s: number, e: { amount: number }) => s + (e.amount || 0), 0);
    const expensesByCategory: Record<string, number> = {};
    for (const e of (expensesRaw || [])) expensesByCategory[e.category] = (expensesByCategory[e.category] || 0) + e.amount;
    const grossProfit = totalRevenue - totalCogs;
    const operatingProfit = grossProfit - totalExpenses;
    const profitLoss = {
      revenue: totalRevenue,
      cogs: totalCogs,
      gross_profit: grossProfit,
      expenses: totalExpenses,
      expenses_by_category: expensesByCategory,
      operating_profit: operatingProfit,
      margin_rate: totalRevenue > 0 ? Math.round((operatingProfit / totalRevenue) * 1000) / 10 : 0,
    };

    return NextResponse.json({
      period: { from: fromDate, to: toDate },
      sales: saleSummary,                  // 오프라인 판매 전체 (RS 포함 — 호환 "오프라인 판매 총액")
      product_sales: productSalesSummary,  // 제품 매출 (RS 제외, 납품 포함, B2C/B2B)
      repair_sales: repairSalesSummary,    // 복원수리 = A(접수) + B(판매RS) + C(납품RS)
      total_revenue: totalRevenue,         // = product_sales.total + repair_sales.total
      purchases: purchaseSummary,
      vat: vatSummary,
      margin,
      by_product: byProduct,
      by_supplier: supplierRanking,
      payables: { items: payables, total: totalPayables },
      receivables: { items: receivables, total: totalReceivables, aging: agingSummary },
      daily: { sales: dailySales, purchases: dailyPurchases, repairs: dailyRepairs, product: dailyProduct },
      profit_loss: profitLoss,
      details: { sales, purchases, repairs: repairSales, deliveries },
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

// 결제수단별 매출(현금주의) — 복합결제는 payment_detail(실수납)로 카드/현금/이체 분해, 비복합은 total_amount 유지
// → 카드사 정산·부가세 카드매출과 정합. (2026-08-06)
function buildByMethod(
  sales: Array<{ payment_method: string; total_amount: number; payment_detail: Record<string, number> | null }>,
): Record<string, number> {
  const result: Record<string, number> = {};
  for (const r of sales) {
    if (r.payment_method === 'mixed' && r.payment_detail) {
      for (const m of ['card', 'cash', 'transfer'] as const) {
        const amt = r.payment_detail[m] || 0;
        if (amt) result[m] = (result[m] || 0) + amt;
      }
    } else {
      const k = r.payment_method || 'unknown';
      result[k] = (result[k] || 0) + (r.total_amount || 0);
    }
  }
  return result;
}
