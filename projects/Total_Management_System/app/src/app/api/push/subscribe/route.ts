import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/push/subscribe — FCM 토큰 등록
 *  같은 user_id의 옛 토큰은 자동 삭제 → single-token-per-user 정책으로
 *  같은 사용자가 캐시 비움/재구독 등으로 다중 토큰을 누적하는 문제 차단.
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { token, deviceInfo } = await req.json();
    if (!token) return NextResponse.json({ error: 'token required' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 1) 같은 user_id의 다른 토큰 삭제 (현재 등록할 token은 유지)
    await db.from('push_subscriptions')
      .delete()
      .eq('user_id', user.id)
      .neq('token', token);

    // 2) 현재 토큰 upsert
    await db.from('push_subscriptions').upsert(
      { user_id: user.id, token, device_info: deviceInfo || null, updated_at: new Date().toISOString() },
      { onConflict: 'token' },
    );

    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
