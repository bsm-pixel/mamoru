import { NextRequest, NextResponse } from 'next/server';
import { cancel } from '@/lib/lotte/client';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/lotte/cancel — 송장 취소 */
export async function POST(request: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { invNo, orderId } = await request.json();
    const result = await cancel(invNo);

    // 성공 시 주문에서 송장 정보 제거
    if (result.success && orderId) {
      await (supabase as any)
        .from('orders')
        .update({
          invoice_number: null,
          courier_code: null,
          courier_name: null,
          status: 'pay_done',
          shipped_at: null,
        })
        .eq('id', orderId);
    }

    return NextResponse.json(result);
  } catch (err) {
    console.error('[lotte/cancel] 송장 취소 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
