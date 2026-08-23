import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateInvoice, prepareImwebDelivery } from '@/lib/imweb/client';

/** POST /api/imweb/push-invoice — 아임웹에 송장번호 재전송 */
export async function POST(request: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { orderId } = await request.json();
    if (!orderId) return NextResponse.json({ error: 'MISSING_ORDER_ID' }, { status: 400 });

    // Supabase에서 주문 정보 조회
    const { data: order } = await (supabase as any)
      .from('orders')
      .select('imweb_order_no, invoice_number')
      .eq('id', orderId)
      .single();

    if (!order?.imweb_order_no || !order?.invoice_number) {
      return NextResponse.json({ error: 'NO_INVOICE' }, { status: 400 });
    }

    // 128: 배송대기 자동 전환(place) 후 송장 등록 — 아임웹 수동 "배송대기 처리" 불필요
    await prepareImwebDelivery(order.imweb_order_no);
    const result = await updateInvoice(order.imweb_order_no, {
      parcel_code: 'LOTTE',
      invoice_no: order.invoice_number,
    });

    return NextResponse.json(result);
  } catch (err) {
    console.error('[imweb/push-invoice] 오류:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
