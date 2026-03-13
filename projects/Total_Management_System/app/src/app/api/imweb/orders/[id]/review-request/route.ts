import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

/** POST /api/imweb/orders/[id]/review-request — 수동 리뷰 요청 알림톡 발송 */
export async function POST(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const supabase: any = createServiceClient();

    // 주문 조회
    const { data: order, error } = await supabase
      .from('orders')
      .select('id, status, orderer_name, orderer_phone, review_requested_at')
      .eq('id', id)
      .single();

    if (error || !order) {
      return NextResponse.json({ error: '주문을 찾을 수 없습니다' }, { status: 404 });
    }

    // 배송완료/구매확정 상태만 허용
    if (!['delivered', 'confirmed'].includes(order.status)) {
      return NextResponse.json({ error: '배송완료 상태의 주문만 리뷰 요청 가능합니다' }, { status: 400 });
    }

    // 이미 발송된 경우
    if (order.review_requested_at) {
      return NextResponse.json({ error: '이미 리뷰 요청이 발송되었습니다', requested_at: order.review_requested_at }, { status: 409 });
    }

    // 주문 상품 목록 조회
    const { data: items } = await supabase
      .from('order_items')
      .select('product_name, quantity')
      .eq('order_id', id);

    const productNames = (items || [])
      .map((it: { product_name: string; quantity: number }) =>
        it.quantity > 1 ? `${it.product_name} x${it.quantity}` : it.product_name
      )
      .join(', ');

    const phone = order.orderer_phone;
    const name = order.orderer_name;

    if (!phone || !name) {
      return NextResponse.json({ error: '주문자 연락처 정보가 없습니다' }, { status: 400 });
    }

    // 알림톡 발송
    const result = await sendNotification({
      template: 'purchase_review_request',
      phone,
      name,
      data: {
        order_uid: id,
        product_names: productNames,
        review_type: 'purchase',
        name,
      },
    });

    if (!result.success) {
      return NextResponse.json({ error: `알림톡 발송 실패: ${result.error}` }, { status: 500 });
    }

    // 발송 기록
    await supabase
      .from('orders')
      .update({ review_requested_at: new Date().toISOString() })
      .eq('id', id);

    return NextResponse.json({ success: true, message: '리뷰 요청 알림톡이 발송되었습니다' });
  } catch (err) {
    console.error('[review-request] 발송 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
