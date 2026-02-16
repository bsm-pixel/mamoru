import { NextResponse } from 'next/server';
import { syncOrders } from '@/lib/imweb/sync';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/imweb/sync — 수동 주문 동기화 */
export async function POST() {
  try {
    // 인증 확인
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const result = await syncOrders();
    return NextResponse.json(result);
  } catch (err) {
    console.error('[sync] 동기화 실패:', err);
    return NextResponse.json(
      { error: String(err) },
      { status: 500 }
    );
  }
}
