import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

/**
 * POST /api/orders/[id]/pickup-complete — 직접수령(대면 픽업) 완료 처리
 *
 * 고객이 매장에서 아임웹으로(쿠폰 등) 결제 후 직접 수령하는 경우.
 * 송장 없이 주문을 delivered 로 마감(+ is_pickup 마커). 재고는 이미 sync 가 차감했으므로 손대지 않는다.
 * 배송완료와 동일하게 리뷰 요청 알림톡을 발송한다(미발송인 경우).
 * 시리얼 배정이 필요하면 상세패널의 '배정 시리얼'로 별도 처리(취소 아님이면 언제든 가능).
 * ⚠️ 아임웹 상태 역동기 불가(code -99) → 아임웹에서도 수동 완료 필요(imwebManual).
 */
export async function POST(_request: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id: orderId } = await params;

  const auth = await createServerSupabaseClient();
  const { data: { user } } = await auth.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db: any = createServiceClient();
  const { data: order } = await db
    .from('orders')
    .select('id, status, orderer_name, orderer_phone, review_requested_at')
    .eq('id', orderId).single();
  if (!order) return NextResponse.json({ error: '주문을 찾을 수 없습니다' }, { status: 404 });
  if (['delivered', 'cancelled'].includes(order.status)) return NextResponse.json({ ok: true, already: true });

  const now = new Date().toISOString();
  await db
    .from('orders')
    .update({ status: 'delivered', delivered_at: now })
    .eq('id', orderId);

  // is_pickup 마커 — 마이그126 미실행 시 컬럼이 없어 실패할 수 있으므로 별도·best-effort (픽업 완료 자체는 위에서 끝남)
  const { error: pickupErr } = await db.from('orders').update({ is_pickup: true }).eq('id', orderId);
  if (pickupErr) console.warn('[pickup] is_pickup 마커 실패(마이그126 미실행?):', pickupErr.message);

  // 리뷰 요청 알림톡 — 배송완료와 동일(제품 받은 시점). 미발송 + 연락처 있을 때만.
  if (!order.review_requested_at && order.orderer_phone && order.orderer_name) {
    const { data: items } = await db.from('order_items').select('product_name, quantity').eq('order_id', orderId);
    const productNames = (items || [])
      .map((it: { product_name: string; quantity: number }) => (it.quantity > 1 ? `${it.product_name} x${it.quantity}` : it.product_name))
      .join(', ');
    await sendNotification({
      template: 'purchase_review_request',
      phone: order.orderer_phone,
      name: order.orderer_name,
      data: { order_uid: orderId, product_names: productNames, review_type: 'purchase', name: order.orderer_name },
    }).catch((e: unknown) => console.error('[pickup] 리뷰 요청 발송 실패:', e));
    await db.from('orders').update({ review_requested_at: now }).eq('id', orderId);
  }

  return NextResponse.json({ ok: true, imwebManual: true });
}
