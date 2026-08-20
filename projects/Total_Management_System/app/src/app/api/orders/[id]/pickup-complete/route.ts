import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

/**
 * POST /api/orders/[id]/pickup-complete — 직접수령(대면 픽업) 완료 처리
 *
 * 고객이 매장에서 아임웹으로(쿠폰 등) 결제 후 직접 수령하는 경우.
 * 송장 없이 주문을 delivered 로 마감한다. 재고는 이미 sync 가 차감했으므로 손대지 않는다.
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
  const { data: order } = await db.from('orders').select('id, status').eq('id', orderId).single();
  if (!order) return NextResponse.json({ error: '주문을 찾을 수 없습니다' }, { status: 404 });
  if (['delivered', 'cancelled'].includes(order.status)) return NextResponse.json({ ok: true, already: true });

  await db
    .from('orders')
    .update({ status: 'delivered', delivered_at: new Date().toISOString() })
    .eq('id', orderId);

  return NextResponse.json({ ok: true, imwebManual: true });
}
