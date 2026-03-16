import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { queryStatus } from '@/lib/lotte/client';

/** GET /api/cron/track-delivery — 배송중 주문 추적 → 배송완료 자동 전환 */
export async function GET(request: NextRequest) {
  const authHeader = request.headers.get('authorization');
  if (authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  try {
    const supabase = await createServerSupabaseClient();

    /* shipping 상태 + 송장번호 있는 주문만 조회 */
    const { data: orders, error } = await (supabase as any)
      .from('orders')
      .select('id, invoice_number, imweb_order_no, orderer_name, review_requested_at')
      .eq('status', 'shipping')
      .not('invoice_number', 'is', null)
      .limit(50);

    if (error) throw error;
    if (!orders || orders.length === 0) {
      return NextResponse.json({ checked: 0, delivered: 0 });
    }

    let deliveredCount = 0;

    for (const order of orders) {
      try {
        const result = await queryStatus(order.invoice_number);

        if (result.ok && result.state === 'DELIVERED') {
          await (supabase as any)
            .from('orders')
            .update({
              status: 'delivered',
              delivered_at: new Date().toISOString(),
            })
            .eq('id', order.id);

          deliveredCount++;
          console.log(`[track-delivery] ${order.imweb_order_no} → 배송완료`);
        }
      } catch (e) {
        console.error(`[track-delivery] ${order.imweb_order_no} 추적 실패:`, e);
      }
    }

    return NextResponse.json({
      checked: orders.length,
      delivered: deliveredCount,
    });
  } catch (err) {
    console.error('[cron/track-delivery] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
