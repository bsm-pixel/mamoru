import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getNextInvoice, bookShipment, cancelShipment } from '@/lib/lotte/alps-client';

/** POST /api/sales/[id]/ship — 판매 건 송장 생성 */
export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 판매 건 조회
    const { data: sale, error: fetchErr } = await db
      .from('offline_sales')
      .select('id, customer_id, customer_name, customer_phone, invoice_number, cancelled_at')
      .eq('id', id)
      .single();

    if (fetchErr || !sale) return NextResponse.json({ error: '판매 건을 찾을 수 없습니다' }, { status: 404 });
    if (sale.cancelled_at) return NextResponse.json({ error: '취소된 판매입니다' }, { status: 400 });
    if (sale.invoice_number) return NextResponse.json({ error: '이미 송장이 생성되었습니다' }, { status: 400 });

    // 고객 주소 조회
    if (!sale.customer_id) return NextResponse.json({ error: '고객 정보가 없어 송장 생성 불가' }, { status: 400 });

    const { data: customer } = await db
      .from('customers')
      .select('postcode, address_road, address_detail')
      .eq('id', sale.customer_id)
      .single();

    if (!customer?.postcode || !customer?.address_road) {
      return NextResponse.json({ error: '고객 주소(우편번호+도로명)를 먼저 등록해주세요' }, { status: 400 });
    }

    // 판매 항목 조회 (품목명용)
    const { data: items } = await db
      .from('offline_sale_items')
      .select('product_name, quantity')
      .eq('sale_id', id);

    const goodsName = items && items.length > 0
      ? items.map((i: { product_name: string; quantity: number }) => `${i.product_name}×${i.quantity}`).join(', ')
      : '마모루 제품';

    // ALPS 송장 생성
    const { invoiceNumber } = await getNextInvoice();
    const fullAddr = [customer.address_road, customer.address_detail].filter(Boolean).join(' ');

    const result = await bookShipment({
      invoiceNumber,
      receiverName: sale.customer_name,
      receiverTel: sale.customer_phone || '',
      receiverZip: customer.postcode,
      receiverAddr: fullAddr,
      goodsName: goodsName.slice(0, 50), // ALPS 50자 제한
    });

    if (!result.success) {
      return NextResponse.json({ error: `ALPS 송장 생성 실패: ${result.error}` }, { status: 502 });
    }

    // 판매 건 업데이트
    await db.from('offline_sales').update({
      invoice_number: invoiceNumber,
      shipped_at: new Date().toISOString(),
      delivery_method: 'shipping',
      courier_name: '롯데택배',
    }).eq('id', id);

    return NextResponse.json({ success: true, invoiceNumber });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/sales/[id]/ship — 송장 취소 */
export async function DELETE(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: sale } = await db
      .from('offline_sales')
      .select('invoice_number')
      .eq('id', id)
      .single();

    if (!sale?.invoice_number) return NextResponse.json({ error: '송장이 없습니다' }, { status: 400 });

    // ALPS 취소 시도
    let warning: string | undefined;
    const cancelResult = await cancelShipment(sale.invoice_number);
    if (!cancelResult.success) {
      warning = `ALPS 취소 실패: ${cancelResult.error}`;
    }

    // DB 업데이트
    await db.from('offline_sales').update({
      invoice_number: null,
      shipped_at: null,
      delivery_method: 'pickup',
    }).eq('id', id);

    return NextResponse.json({ success: true, warning });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
