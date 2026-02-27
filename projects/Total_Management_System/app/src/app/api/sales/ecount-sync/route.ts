import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { saveSale } from '@/lib/ecount/sales';
import { isStatusOk } from '@/lib/ecount/client';

/** POST /api/sales/ecount-sync — 오프라인 판매 → 이카운트 전표 동기화 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { saleId } = await req.json() as { saleId: string };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 판매 + 항목 조회
    const [saleRes, itemsRes] = await Promise.all([
      db.from('offline_sales').select('*').eq('id', saleId).single(),
      db.from('offline_sale_items').select('*').eq('sale_id', saleId),
    ]);

    if (saleRes.error) throw saleRes.error;
    const sale = saleRes.data;
    const items = itemsRes.data || [];

    // 이카운트 거래처 코드 조회
    let customerCode = '';
    if (sale.customer_id) {
      const { data: cust } = await db
        .from('customers')
        .select('ecount_customer_code')
        .eq('id', sale.customer_id)
        .single();
      customerCode = cust?.ecount_customer_code || '';
    }

    // 이카운트 판매전표 생성
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ecountItems = items.map((item: any) => ({
      PROD_CD: item.sku || '',
      PROD_DES: item.product_name,
      QTY: String(item.quantity),
      PRICE: String(item.unit_price),
      SUPPLY_AMT: String(item.total_price),
    }));

    const result = await saveSale({
      saleDate: sale.sale_date,
      customerCode,
      items: ecountItems,
      remarks: `TMS ${sale.sale_number}`,
    });

    if (isStatusOk(result.Status)) {
      await db
        .from('offline_sales')
        .update({
          ecount_sync_status: 'synced',
          ecount_synced_at: new Date().toISOString(),
        })
        .eq('id', saleId);

      return NextResponse.json({ success: true, ecount: result.Data });
    } else {
      await db
        .from('offline_sales')
        .update({ ecount_sync_status: 'failed' })
        .eq('id', saleId);

      return NextResponse.json(
        { error: `이카운트 동기화 실패: ${result.Error?.Message || 'Unknown'}` },
        { status: 502 },
      );
    }
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
