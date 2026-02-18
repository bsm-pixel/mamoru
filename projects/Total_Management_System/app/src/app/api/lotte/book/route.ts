import { NextRequest, NextResponse } from 'next/server';
import { bookSingle } from '@/lib/lotte/client';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateInvoice } from '@/lib/imweb/client';

/** POST /api/lotte/book — 송장 생성 */
export async function POST(request: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await request.json();

    // 상품명 자동 조합: order_items에서 가져오기
    if (!body.gdsNm && body.orderId) {
      const { data: items } = await (supabase as any)
        .from('order_items')
        .select('product_name, quantity')
        .eq('order_id', body.orderId);
      if (items && items.length > 0) {
        body.gdsNm = items
          .map((it: { product_name: string; quantity: number }) =>
            it.quantity > 1 ? `${it.product_name} x${it.quantity}` : it.product_name
          )
          .join(', ')
          .slice(0, 750); // ALPS gdsNm 최대 750byte
      }
    }

    const result = await bookSingle(body);

    // 주문에 송장 정보 저장
    if (body.orderId) {
      await (supabase as any)
        .from('orders')
        .update({
          invoice_number: result.invNo,
          courier_code: 'LOTTE',
          courier_name: '롯데택배',
          status: 'shipping',
          shipped_at: new Date().toISOString(),
        })
        .eq('id', body.orderId);
    }

    // 아임웹에 송장번호 반영 (배송대기 상태일 때만 성공)
    let imwebSynced = false;
    let imwebNeedsManual = false;
    if (body.ordNo) {
      try {
        const imwebResult = await updateInvoice(body.ordNo, {
          parcel_code: 'LOTTE',
          invoice_no: result.invNo,
        });
        imwebSynced = imwebResult.success;
        imwebNeedsManual = imwebResult.needsManual;
      } catch (imwebErr) {
        console.warn('[lotte/book] 아임웹 반영 실패:', imwebErr);
        imwebNeedsManual = true;
      }
    }

    return NextResponse.json({
      ...result,
      imwebSynced,
      imwebNeedsManual,
    });
  } catch (err) {
    console.error('[lotte/book] 송장 생성 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
