import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { queryStatus } from '@/lib/lotte/client';

/** POST /api/lotte/check-cancel — ALPS 추적으로 취소 여부 확인 후 자동 전환 */
export async function POST(request: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { orderId, invNo } = await request.json();
    if (!orderId || !invNo) {
      return NextResponse.json({ error: 'MISSING_PARAMS' }, { status: 400 });
    }

    // ALPS 추적 API로 상태 확인
    const track = await queryStatus(invNo);
    console.log('[check-cancel] 추적 결과:', { invNo, state: track.state });

    if (track.state === 'CANCELLED' || track.state === 'NOT_FOUND') {
      // Supabase 상태를 cancelled로 전환
      await (supabase as any)
        .from('orders')
        .update({
          status: 'cancelled',
          invoice_number: null,
          courier_code: null,
          courier_name: null,
          shipped_at: null,
        })
        .eq('id', orderId);

      console.log('[check-cancel] 자동 취소 전환 완료:', orderId);
      // 아임웹 상태 변경은 v2 API 미지원 → 수동 처리 필요
      return NextResponse.json({
        cancelled: true,
        state: track.state,
        imwebManual: true, // 아임웹에서 수동 취소 필요
      });
    }

    // 아직 취소되지 않음
    return NextResponse.json({ cancelled: false, state: track.state });
  } catch (err) {
    console.error('[check-cancel] 오류:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
