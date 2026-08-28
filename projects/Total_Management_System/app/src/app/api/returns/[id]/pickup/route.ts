import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getNextInvoice, bookReturnPickup } from '@/lib/lotte/alps-client';

/** POST /api/returns/[id]/pickup — 롯데 반품 수거접수 (ustRtgSctCd='02', 고객집 → 마모루 회수)
 *  롯데 IS팀 회신(2026-08-27): 02=반품, 양식은 출고와 동일, orglInvNo=원송장(선택), 취소 API 미지원.
 *  ⚠️ 접수 후 취소는 ALPS 화면에서 수동(취소 API 없음).
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
    if (r.pickup_invoice_number) {
      return NextResponse.json({ error: '이미 반품 수거접수가 되어 있습니다' }, { status: 400 });
    }
    if (r.status === 'cancelled' || r.status === 'completed') {
      return NextResponse.json({ error: '취소·완료된 건은 수거접수할 수 없습니다' }, { status: 400 });
    }

    // 고객(수거지) 주소 — 고객 정보에서 조회(교환 출고와 동일 패턴)
    let custName = r.name as string | null;
    let custTel = r.phone as string | null;
    let postcode: string | null = null;
    let addrRoad: string | null = null;
    let addrDetail: string | null = null;
    if (r.customer_id) {
      const { data: c } = await db.from('customers').select('name, phone, postcode, address_road, address_detail').eq('id', r.customer_id).single();
      if (c) {
        custName = custName || c.name;
        custTel = custTel || c.phone;
        postcode = c.postcode; addrRoad = c.address_road; addrDetail = c.address_detail;
      }
    }
    if (!postcode || !addrRoad) {
      return NextResponse.json({ error: '고객 주소(우편번호+도로명)가 없습니다. 고객 정보에서 보강 후 다시 시도해주세요.' }, { status: 400 });
    }
    if (!custTel) {
      return NextResponse.json({ error: '고객 연락처가 없습니다.' }, { status: 400 });
    }

    const goodsName = `${r.product_name || '반품 상품'}${r.serial_number ? ` (${r.serial_number})` : ''} 회수`.slice(0, 50);
    const fullAddr = [addrRoad, addrDetail].filter(Boolean).join(' ');

    // 반품 수거접수 — 보내는분(수거지)=고객, 받는분=마모루, ustRtgSctCd='02'. 원송장은 있으면 참조(선택)
    const { invoiceNumber } = await getNextInvoice();
    const result = await bookReturnPickup({
      invoiceNumber,
      pickupName: custName || '고객',
      pickupTel: custTel,
      pickupZip: postcode,
      pickupAddr: fullAddr,
      goodsName,
      deliveryMessage: '반품 수거',
      orgInvoiceNumber: r.exchange_out_invoice_number || undefined,  // 원송장(있으면, 선택)
    });
    if (!result.success) {
      return NextResponse.json({ error: `롯데 반품 수거접수 실패: ${result.error}` }, { status: 502 });
    }

    // 송장 저장 + 상태 '수거예약'으로 전이
    const { data, error } = await db.from('returns').update({
      pickup_invoice_number: invoiceNumber,
      pickup_courier_name: '롯데택배',
      pickup_booked_at: new Date().toISOString(),
      status: 'pickup_scheduled',
      pickup_scheduled_at: new Date().toISOString(),
    }).eq('id', id).select().single();
    if (error) {
      console.error('[returns pickup] DB 저장 실패(ALPS는 성공):', invoiceNumber, error);
      return NextResponse.json({ success: true, warning: 'DB 저장 실패 — 송장은 발급됨(ALPS 확인)', invoiceNumber });
    }
    return NextResponse.json({ success: true, return: data });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
