import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { restoreOrderStock } from '@/lib/imweb/sync';

/**
 * POST /api/orders/[id]/cancel — TMS에서 주문 취소 (재고·시리얼 즉시 복구)
 *
 * 기존 useCancelOrder 는 status 만 바꾸고 재고/시리얼 복구는 아임웹 취소 동기화 때로 미뤘다.
 * 이 라우트는 취소 즉시 restoreOrderStock 으로 재고·배정 시리얼을 복구한다(멱등: stock_deducted 플래그).
 * ⚠️ 아임웹 주문 상태 역동기는 불가(code -99) → 아임웹에서도 수동 취소 필요(imwebManual).
 */
export async function POST(_request: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id: orderId } = await params;

  const auth = await createServerSupabaseClient();
  const { data: { user } } = await auth.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db: any = createServiceClient();
  const { data: order } = await db.from('orders').select('id, status, stock_deducted').eq('id', orderId).single();
  if (!order) return NextResponse.json({ error: '주문을 찾을 수 없습니다' }, { status: 404 });
  if (order.status === 'cancelled') return NextResponse.json({ ok: true, already: true });

  // 이미 차감된 경우에만 복구 (멱등). restore 가 배정 시리얼 in_stock 복원 + raw 보정 + 아임웹 재고 +qty.
  if (order.stock_deducted) {
    await restoreOrderStock(orderId);
    await db.from('orders').update({ stock_deducted: false }).eq('id', orderId);
  }
  await db.from('orders').update({ status: 'cancelled' }).eq('id', orderId);

  return NextResponse.json({ ok: true, imwebManual: true });
}
