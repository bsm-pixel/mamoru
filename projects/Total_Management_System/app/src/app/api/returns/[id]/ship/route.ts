import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getNextInvoice, bookShipment } from '@/lib/lotte/alps-client';

/** POST /api/returns/[id]/ship — 교환 출고 송장 발행 (새 제품 1개만, 매장→고객 정방향)
 *  빠른송장(manual-invoices)과 동일한 롯데 발행 패턴. 판매/주문/아임웹 부수효과 0.
 */
export async function POST(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: r } = await db.from('returns').select('*').eq('id', id).single();
    if (!r) return NextResponse.json({ error: '반품 건을 찾을 수 없습니다' }, { status: 404 });
    if (r.return_type !== 'exchange' || !r.new_product_name) {
      return NextResponse.json({ error: '교환(새 제품) 건이 아닙니다' }, { status: 400 });
    }
    if (r.exchange_out_invoice_number) {
      return NextResponse.json({ error: '이미 교환 출고 송장이 발행되었습니다' }, { status: 400 });
    }

    // 받는분 주소 — 고객 정보에서 조회(빠른송장과 동일)
    let receiverName = r.name as string | null;
    let receiverTel = r.phone as string | null;
    let postcode: string | null = null;
    let addrRoad: string | null = null;
    let addrDetail: string | null = null;
    if (r.customer_id) {
      const { data: c } = await db.from('customers').select('name, phone, postcode, address_road, address_detail').eq('id', r.customer_id).single();
      if (c) {
        receiverName = receiverName || c.name;
        receiverTel = receiverTel || c.phone;
        postcode = c.postcode; addrRoad = c.address_road; addrDetail = c.address_detail;
      }
    }
    // 주문 교환 등 고객 미연결 건 — returns 레코드 자체 주소(수령지) 사용
    if (!postcode || !addrRoad) {
      postcode = postcode || (r.postcode as string | null);
      addrRoad = addrRoad || (r.address as string | null);
      addrDetail = addrDetail || (r.address_detail as string | null);
    }
    if (!postcode || !addrRoad) {
      return NextResponse.json({ error: '수령지 주소(우편번호+주소)가 없습니다.' }, { status: 400 });
    }
    if (!receiverTel) {
      return NextResponse.json({ error: '고객 연락처가 없습니다.' }, { status: 400 });
    }

    const goodsName = `${r.new_product_name} ×1 (교환)`.slice(0, 50);
    const fullAddr = [addrRoad, addrDetail].filter(Boolean).join(' ');

    const { invoiceNumber } = await getNextInvoice();
    const result = await bookShipment({
      invoiceNumber,
      receiverName: receiverName || '고객',
      receiverTel,
      receiverZip: postcode,
      receiverAddr: fullAddr,
      goodsName,
      deliveryMessage: '교환 상품',
    });
    if (!result.success) {
      return NextResponse.json({ error: `ALPS 송장 생성 실패: ${result.error}` }, { status: 502 });
    }

    const { data, error } = await db.from('returns').update({
      exchange_out_invoice_number: invoiceNumber,
      exchange_out_courier_name: '롯데택배',
      exchange_shipped_at: new Date().toISOString(),
    }).eq('id', id).select().single();
    if (error) {
      console.error('[returns ship] DB 저장 실패(ALPS는 성공):', invoiceNumber, error);
      return NextResponse.json({ success: true, warning: 'DB 저장 실패 — 송장번호는 발급됨', invoiceNumber });
    }
    // 주문 교환건이면 원 주문에도 송장 미러(주문/판매 화면과 동기)
    if (r.order_id) {
      await db.from('orders').update({ exchange_invoice_number: invoiceNumber, exchange_shipped_at: new Date().toISOString() }).eq('id', r.order_id);
    }
    return NextResponse.json({ success: true, return: data });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
