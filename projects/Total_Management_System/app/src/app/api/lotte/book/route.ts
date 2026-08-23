import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateInvoice, prepareImwebDelivery } from '@/lib/imweb/client';
import { getNextInvoice, bookShipment } from '@/lib/lotte/alps-client';

/** POST /api/lotte/book — 송장 생성 (ALPS 직접 호출) */
export async function POST(request: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await request.json();

    // 상품명 자동 조합: order_items에서 가져오기
    if (!body.gdsNm && body.orderId) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
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
          .slice(0, 750);
      }
    }

    // ALPS 직접 호출 — 송장번호 발급 + 접수
    const { invoiceNumber } = await getNextInvoice();

    const result = await bookShipment({
      invoiceNumber,
      receiverName: body.rcvName || '',
      receiverTel: body.rcvTel || '',
      receiverZip: body.rcvZip || '',
      receiverAddr: body.rcvAdr || '',
      goodsName: body.gdsNm || '마모루 제품',
      deliveryMessage: body.dlvMsg || '',
    });

    if (!result.success) {
      return NextResponse.json(
        { error: `ALPS 송장 생성 실패: ${result.error}` },
        { status: 502 }
      );
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 주문에 송장 정보 저장
    // 🔴 128 (2026-08-23): 송장 발급 ≠ 출고. status 를 'ready_to_ship'(배송대기)로 두고
    //    shipped_at 은 비워둔다. 기사 집하가 감지되면 크론[1-A]이 'shipping'(배송중)+shipped_at 을 채운다.
    //    (복원수리·납품과 동일 리듬 — 송장만으로 "배송중"이 뜨던 것을 바로잡음)
    if (body.orderId) {
      await db
        .from('orders')
        .update({
          invoice_number: invoiceNumber,
          courier_code: 'LOTTE',
          courier_name: '롯데택배',
          status: 'ready_to_ship',
        })
        .eq('id', body.orderId);
    }

    // 납품(B2B)에 송장 정보 저장
    // 🔴 110 (2026-07-12): 여기서 status='shipped' 를 강제하던 것을 제거.
    //    송장 발급 ≠ 출고다. 기사님이 아직 안 왔는데 화면에 "출고완료"가 뜨고 있었다(사장님 지적).
    //    → status 는 'confirmed'(출고대기) 로 두고, 크론 [4-A] 집하 감지가 'shipped' 로 올린다.
    //    B2C 판매(api/sales/[id]/ship)와 같은 정의로 통일 — 그쪽도 송장번호만 넣고 출고는 집하가 채운다.
    if (body.deliveryId) {
      await db
        .from('deliveries')
        .update({
          tracking_number: invoiceNumber,
          updated_at: new Date().toISOString(),
        })
        .eq('id', body.deliveryId);
    }

    // 아임웹 역동기: 배송대기 전환(place) → 송장번호 등록(invoice)
    // 🔴 128: 이제 아임웹 "배송대기 처리"를 자동으로 수행하므로 수동 개입 불필요.
    //    place 로 전 품목을 배송대기(STANDBY)로 올린 뒤 invoice 로 송장번호를 등록한다.
    //    (배송중 전환은 집하 시 크론이 send 로 수행 — 여기서는 배송대기까지만)
    let imwebSynced = false;
    let imwebNeedsManual = false;
    if (body.ordNo) {
      try {
        await prepareImwebDelivery(body.ordNo); // 결제완료 → 배송대기 (전 품목)
        const imwebResult = await updateInvoice(body.ordNo, {
          parcel_code: 'LOTTE',
          invoice_no: invoiceNumber,
        });
        imwebSynced = imwebResult.success;
        imwebNeedsManual = imwebResult.needsManual;
      } catch (imwebErr) {
        console.warn('[lotte/book] 아임웹 반영 실패:', imwebErr);
        imwebNeedsManual = true;
      }
    }

    return NextResponse.json({
      ok: true,
      invNo: invoiceNumber,
      imwebSynced,
      imwebNeedsManual,
    });
  } catch (err) {
    console.error('[lotte/book] 송장 생성 실패:', err);
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}
