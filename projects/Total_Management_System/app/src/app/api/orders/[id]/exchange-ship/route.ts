import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { getNextInvoice, bookShipment } from '@/lib/lotte/alps-client';
import { sendNotification } from '@/lib/notification/make-webhook';

/**
 * POST /api/orders/[id]/exchange-ship — 교환 새 제품 발송 송장 발행 (롯데, 수령지 기준)
 *   원 주문 invoice_number/status 는 원본 배송 그대로. 교환품은 별도 송장(exchange_invoice_number)으로 추적.
 *   교환 처리 시 발송방식='배송'인 주문에서만 사용.
 */
export async function POST(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id: orderId } = await params;
    const auth = await createServerSupabaseClient();
    const { data: { user } } = await auth.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db: any = createServiceClient();
    const { data: o } = await db.from('orders')
      .select('id, exchanged_at, exchange_goods, exchange_invoice_number, recipient_name, orderer_name, recipient_phone, orderer_phone, recipient_postcode, recipient_address, recipient_address_detail')
      .eq('id', orderId).single();
    if (!o) return NextResponse.json({ error: '주문을 찾을 수 없습니다' }, { status: 404 });
    if (!o.exchanged_at) return NextResponse.json({ error: '교환 처리된 주문이 아닙니다' }, { status: 400 });
    if (o.exchange_invoice_number) return NextResponse.json({ error: '이미 교환품 송장이 발행되었습니다' }, { status: 400 });

    const receiverName = o.recipient_name || o.orderer_name || '고객';
    const receiverTel = o.recipient_phone || o.orderer_phone;
    const postcode = o.recipient_postcode;
    const addr = [o.recipient_address, o.recipient_address_detail].filter(Boolean).join(' ');
    if (!postcode || !o.recipient_address) return NextResponse.json({ error: '수령지 주소(우편번호+주소)가 없습니다' }, { status: 400 });
    if (!receiverTel) return NextResponse.json({ error: '수령인 연락처가 없습니다' }, { status: 400 });

    const goodsName = `${o.exchange_goods || '교환 상품'} (교환)`.slice(0, 50);

    const { invoiceNumber } = await getNextInvoice();
    const result = await bookShipment({
      invoiceNumber,
      receiverName,
      receiverTel,
      receiverZip: postcode,
      receiverAddr: addr,
      goodsName,
      deliveryMessage: '교환 상품',
    });
    if (!result.success) {
      return NextResponse.json({ error: `롯데 송장 생성 실패: ${result.error}` }, { status: 502 });
    }

    const now = new Date().toISOString();
    const { error } = await db.from('orders').update({
      exchange_invoice_number: invoiceNumber,
      exchange_shipped_at: now,
      updated_at: now,
    }).eq('id', orderId);
    if (error) {
      console.error('[orders/exchange-ship] DB 저장 실패(ALPS는 성공):', invoiceNumber, error);
      return NextResponse.json({ success: true, warning: 'DB 저장 실패 — 송장은 발급됨(ALPS 확인)', invoiceNumber });
    }

    /* 고객 알림톡 — 교환품 발송 안내 (2026-09-21 신규)
       🚨 교환 출고 경로가 **두 곳**이다. 둘 다 넣지 않으면 한쪽만 조용히 무발송이 된다.
          · 여기(주문 화면 [교환품 송장])  · `/api/returns/[id]/ship`(반품·교환 화면)
       위 update 성공 뒤(= 송장번호 확정)에만 부르고, 이미 발행된 건은 위에서 400 으로 끊기므로 건당 1회. */
    if (receiverTel) {
      after(async () => {
        try {
          await sendNotification({
            template: 'exchange_shipped',
            phone: String(receiverTel),
            name: String(receiverName || ''),
            data: {
              // 🚨 셋 중 하나라도 비면 문자 대체 + 버튼 소멸 → 폴백을 둔다
              product_name: String(o.exchange_goods || '교환 상품'),
              courier: '롯데택배',          // ALPS 발행분은 항상 롯데
              tracking: invoiceNumber,
            },
          });
        } catch (e) { console.error('[orders/exchange-ship] 교환발송 알림톡 실패(송장은 발행됨):', invoiceNumber, e); }
      });
    }

    return NextResponse.json({ success: true, invoiceNumber });
  } catch (err) {
    console.error('[orders/exchange-ship]', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
