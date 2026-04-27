import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getNextInvoice, bookShipment } from '@/lib/lotte/alps-client';

/** POST /api/manual-invoices — 빠른 송장(별도 송장) 발급
 *  body: { customer_id: string; goods_name: string; delivery_message?: string }
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json().catch(() => ({}));
    const customerId: string | undefined = body.customer_id;
    const goodsNameRaw: string | undefined = body.goods_name;
    const deliveryMessage: string | null = (body.delivery_message ?? '').trim() || null;

    if (!customerId) {
      return NextResponse.json({ error: '고객을 선택해주세요' }, { status: 400 });
    }
    const goodsName = (goodsNameRaw ?? '').trim();
    if (!goodsName) {
      return NextResponse.json({ error: '품목명을 입력해주세요' }, { status: 400 });
    }
    if (goodsName.length > 50) {
      return NextResponse.json({ error: '품목명은 50자 이하여야 합니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 고객 조회
    const { data: customer, error: customerErr } = await db
      .from('customers')
      .select('id, name, phone, postcode, address_road, address_detail')
      .eq('id', customerId)
      .single();

    if (customerErr || !customer) {
      return NextResponse.json({ error: '고객을 찾을 수 없습니다' }, { status: 404 });
    }

    // 안전장치: 주소·연락처 미보유 차단
    if (!customer.postcode || !customer.address_road) {
      return NextResponse.json({
        error: '고객 주소(우편번호+도로명)가 없습니다. 고객 정보 페이지에서 보강 후 다시 시도해주세요.',
        customer_id: customer.id,
        missing: 'address',
      }, { status: 400 });
    }
    if (!customer.phone) {
      return NextResponse.json({
        error: '고객 연락처가 없습니다. 고객 정보 페이지에서 보강 후 다시 시도해주세요.',
        customer_id: customer.id,
        missing: 'phone',
      }, { status: 400 });
    }

    // ALPS 송장 발급
    const { invoiceNumber } = await getNextInvoice();
    const fullAddr = [customer.address_road, customer.address_detail].filter(Boolean).join(' ');

    const result = await bookShipment({
      invoiceNumber,
      receiverName: customer.name,
      receiverTel: customer.phone,
      receiverZip: customer.postcode,
      receiverAddr: fullAddr,
      goodsName,
      deliveryMessage: deliveryMessage ?? undefined,
    });

    if (!result.success) {
      return NextResponse.json({ error: `ALPS 송장 생성 실패: ${result.error}` }, { status: 502 });
    }

    // DB insert (스냅샷)
    const { data: inserted, error: insertErr } = await db
      .from('manual_invoices')
      .insert({
        invoice_number: invoiceNumber,
        customer_id: customer.id,
        customer_name: customer.name,
        customer_phone: customer.phone,
        receiver_postcode: customer.postcode,
        receiver_address_road: customer.address_road,
        receiver_address_detail: customer.address_detail ?? null,
        goods_name: goodsName,
        delivery_message: deliveryMessage,
        created_by: user.id,
      })
      .select()
      .single();

    if (insertErr) {
      // ALPS는 성공했지만 DB insert 실패 — 송장번호는 이미 외부에 발급됨. 로그로 추적.
      console.error('[manual-invoices] DB insert 실패 (ALPS는 성공):', invoiceNumber, insertErr);
      return NextResponse.json({
        success: true,
        warning: 'DB 저장 실패. 송장번호는 발급되었으나 이력에 기록되지 않았습니다.',
        invoiceNumber,
      });
    }

    return NextResponse.json({ success: true, invoice: inserted });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** GET /api/manual-invoices
 *  ?date=today    오늘 활성 건만 (default)
 *  ?customer_id=  특정 고객의 이력 (취소 포함, 최신 20건)
 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const url = new URL(req.url);
    const date = url.searchParams.get('date');
    const customerId = url.searchParams.get('customer_id');

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    let query = db.from('manual_invoices').select('*');

    if (customerId) {
      query = query.eq('customer_id', customerId).order('created_at', { ascending: false }).limit(20);
    } else if (date === 'today' || !date) {
      // KST 기준 오늘 시작 시각
      const kstDate = new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Seoul' }); // YYYY-MM-DD
      const kstStart = `${kstDate}T00:00:00+09:00`;
      query = query
        .gte('created_at', kstStart)
        .is('cancelled_at', null)
        .order('created_at', { ascending: false });
    } else {
      query = query.order('created_at', { ascending: false }).limit(50);
    }

    const { data, error } = await query;
    if (error) return NextResponse.json({ error: error.message }, { status: 500 });

    return NextResponse.json({ invoices: data ?? [] });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
