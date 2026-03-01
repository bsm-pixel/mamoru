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

    if (type === 'purchases') {
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
