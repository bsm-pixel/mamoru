import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { createExcelBuffer } from '@/lib/utils/excel';

/**
 * GET /api/reports/export — 매출/매입 엑셀 다운로드
 * ?type=sales|purchases&from=2026-01-01&to=2026-01-31
 */
export async function GET(req: NextRequest) {
  const supabase = createServiceClient();
  const sp = req.nextUrl.searchParams;
  const type = sp.get('type') || 'sales';
  const fromDate = sp.get('from') || new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString().slice(0, 10);
  const toDate = sp.get('to') || new Date().toISOString().slice(0, 10);

  try {
    let buffer: Buffer;
    let filename: string;

    if (type === 'margin') {
      // 마진 분석 엑셀 — 판매 품목별 매출/원가/이익
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data: sales } = await (supabase as any)
        .from('offline_sales')
        .select('id')
        .gte('sale_date', fromDate)
        .lte('sale_date', toDate);

      const saleIds = (sales || []).map((s: { id: string }) => s.id);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data: saleItems } = saleIds.length > 0
        ? await (supabase as any)
            .from('offline_sale_items')
            .select('product_id, product_name, sku, quantity, unit_price, total_price')
            .in('sale_id', saleIds)
        : { data: [] };

      const productIds = [...new Set((saleItems || []).map((i: { product_id: string }) => i.product_id).filter(Boolean))];
      const purchasePriceMap: Record<string, number> = {};
      if (productIds.length > 0) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const { data: products } = await (supabase as any)
          .from('products')
          .select('id, price_purchase')
          .in('id', productIds);
        for (const p of (products || [])) purchasePriceMap[p.id] = p.price_purchase || 0;
      }

      const rows = (saleItems || []).map((item: { product_id: string; product_name: string; sku: string; quantity: number; unit_price: number; total_price: number }) => {
        const purchasePrice = item.product_id ? (purchasePriceMap[item.product_id] || 0) : 0;
        const cogs = purchasePrice * item.quantity;
        const revenue = item.total_price || (item.unit_price * item.quantity);
        const profit = revenue - cogs;
        return {
          product_name: item.product_name,
          sku: item.sku || '',
          quantity: item.quantity,
          unit_price: item.unit_price,
          revenue,
          purchase_price: purchasePrice,
          cogs,
          profit,
          margin_rate: revenue > 0 ? `${Math.round((profit / revenue) * 1000) / 10}%` : '0%',
        };
      });

      buffer = createExcelBuffer('마진분석', [
        { header: '제품명', key: 'product_name', width: 22 },
        { header: 'SKU', key: 'sku', width: 16 },
        { header: '수량', key: 'quantity', width: 8 },
        { header: '판매단가', key: 'unit_price', width: 14 },
        { header: '매출', key: 'revenue', width: 14 },
        { header: '매입단가', key: 'purchase_price', width: 14 },
        { header: '매출원가', key: 'cogs', width: 14 },
        { header: '이익', key: 'profit', width: 14 },
        { header: '이익률', key: 'margin_rate', width: 10 },
      ], rows);

      filename = `마진분석_${fromDate}_${toDate}.xlsx`;
    } else if (type === 'purchases') {
      // 매입 내역
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data } = await (supabase as any)
        .from('purchase_orders')
        .select('po_number, supplier_name, order_date, total_amount, supply_amount, vat_amount, deposit_amount, balance_amount, status')
        .gte('order_date', fromDate)
        .lte('order_date', toDate)
        .order('order_date', { ascending: false });

      const STATUS_LABEL: Record<string, string> = {
        draft: '작성중', ordered: '발주완료', deposit_paid: '선납완료',
        received: '입고완료', balance_paid: '잔금완료', cancelled: '취소',
      };

      buffer = createExcelBuffer('매입내역', [
        { header: '발주번호', key: 'po_number', width: 20 },
        { header: '매입처', key: 'supplier_name', width: 18 },
        { header: '발주일', key: 'order_date', width: 12 },
        { header: '합계', key: 'total_amount', width: 14 },
        { header: '공급가액', key: 'supply_amount', width: 14 },
        { header: '부가세', key: 'vat_amount', width: 12 },
        { header: '선납금', key: 'deposit_amount', width: 14 },
        { header: '잔금', key: 'balance_amount', width: 14 },
        { header: '상태', key: 'status_label', width: 10 },
      ], (data || []).map((r: Record<string, unknown>) => ({
        ...r,
        status_label: STATUS_LABEL[(r.status as string)] || r.status,
      })));

      filename = `매입내역_${fromDate}_${toDate}.xlsx`;
    } else {
      // 매출 내역
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data } = await (supabase as any)
        .from('offline_sales')
        .select('sale_number, customer_name, sale_date, total_amount, supply_amount, vat_amount, discount_amount, paid_amount, payment_method, payment_status')
        .gte('sale_date', fromDate)
        .lte('sale_date', toDate)
        .order('sale_date', { ascending: false });

      const METHOD_LABEL: Record<string, string> = {
        card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합',
      };
      const STATUS_LABEL: Record<string, string> = {
        paid: '결제완료', unpaid: '미결제', partial: '부분결제', refunded: '환불',
      };

      buffer = createExcelBuffer('매출내역', [
        { header: '판매번호', key: 'sale_number', width: 20 },
        { header: '고객명', key: 'customer_name', width: 14 },
        { header: '판매일', key: 'sale_date', width: 12 },
        { header: '합계', key: 'total_amount', width: 14 },
        { header: '공급가액', key: 'supply_amount', width: 14 },
        { header: '부가세', key: 'vat_amount', width: 12 },
        { header: '할인', key: 'discount_amount', width: 12 },
        { header: '수금액', key: 'paid_amount', width: 14 },
        { header: '결제방식', key: 'method_label', width: 10 },
        { header: '결제상태', key: 'status_label', width: 10 },
      ], (data || []).map((r: Record<string, unknown>) => ({
        ...r,
        method_label: METHOD_LABEL[(r.payment_method as string)] || r.payment_method,
        status_label: STATUS_LABEL[(r.payment_status as string)] || r.payment_status,
      })));

      filename = `매출내역_${fromDate}_${toDate}.xlsx`;
    }

    return new NextResponse(new Uint8Array(buffer), {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(filename)}"`,
      },
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
