import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getNextInvoice, bookShipment, cancelShipment } from '@/lib/lotte/alps-client';

/**
 * POST /api/sales/[id]/ship — 판매 건 송장 등록
 *
 *   기본       : 롯데 ALPS 송장 자동 생성 (주소·연락처 필요)
 *   manual:true: 150 신규 — **다른 택배사로 직접 보낸 경우** 택배사+송장번호만 기록.
 *                ALPS 를 호출하지 않으므로 주소가 없어도 되고, 크론 자동추적 대상에서도 빠진다.
 *                (롯데 집하취소 후 우체국으로 보냈는데 기록할 방법이 없던 문제 — 사장님 2026-09-15)
 */
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

    // ── 150: 타 택배사 직접 발송 — ALPS 미호출, 기록만 ──
    const body = await req.json().catch(() => ({}));
    if (body?.manual === true) {
      const invoiceNumber = String(body.invoice_number || '').trim();
      const courierName = String(body.courier_name || '').trim();
      if (!invoiceNumber) return NextResponse.json({ error: '송장번호를 입력해주세요' }, { status: 400 });
      if (!courierName) return NextResponse.json({ error: '택배사를 선택해주세요' }, { status: 400 });

      await db.from('offline_sales').update({
        invoice_number: invoiceNumber,
        delivery_method: 'shipping',
        courier_name: courierName,
      }).eq('id', id);

      return NextResponse.json({ success: true, invoiceNumber, manual: true });
    }

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

    if (!sale.customer_phone) {
      return NextResponse.json({ error: '고객 연락처가 없습니다. 고객 정보에서 연락처를 입력해주세요.' }, { status: 400 });
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

    // 판매 건 업데이트 (shipped_at은 출고완료 시 별도 설정)
    await db.from('offline_sales').update({
      invoice_number: invoiceNumber,
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
      // ALPS 집하취소는 API 미지원 — DB 송장만 정리하고, 사장님께 ALPS 수동취소 안내
      warning = 'DB 송장은 정리됐습니다 · ⚠️ ALPS에서 직접 집하취소 해주세요';
    }

    // DB 업데이트 (109: 집하 자동감지 흔적도 함께 초기화 — 재출고 시 상태 오염 방지)
    await db.from('offline_sales').update({
      invoice_number: null,
      courier_name: null,        // 150: 택배사도 함께 정리 — 안 지우면 다음 발송에 옛 택배사가 따라붙는다
      shipped_at: null,
      shipped_source: null,
      shipped_notified_at: null,
      delivery_method: 'pickup',
    }).eq('id', id);

    return NextResponse.json({ success: true, warning });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
