import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/lotte/cancel — 소프트 취소 (ALPS 직접 취소 필요) */
export async function POST(request: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { invNo, orderId } = await request.json();
    if (!orderId) return NextResponse.json({ error: 'MISSING_ORDER_ID' }, { status: 400 });

    // Supabase 상태를 cancel_pending으로 변경 (송장번호는 유지 — ALPS 확인용)
    await (supabase as any)
      .from('orders')
      .update({ status: 'cancel_pending' })
      .eq('id', orderId);

    console.log('[lotte/cancel] 소프트 취소:', { invNo, orderId });

    return NextResponse.json({
      success: true,
      via: 'soft_cancel',
      message: 'ALPS에서 직접 집하취소 해주세요',
    });
  } catch (err) {
    console.error('[lotte/cancel] 취소 처리 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
