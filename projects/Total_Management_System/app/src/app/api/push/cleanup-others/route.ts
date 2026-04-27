import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/push/cleanup-others — 현재 토큰만 남기고 본인의 다른 토큰 모두 삭제
 *  body: { token: string }   ← 클라이언트가 현재 발급받은 FCM 토큰
 *
 *  사용처: 설정 → 알림 → "이 기기만 알림 받기" 버튼.
 *  중복 알림 발생 시 사장님이 한 번 클릭으로 본인의 다른 디바이스/캐시 토큰 일괄 정리.
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { token } = await req.json().catch(() => ({}));
    if (!token || typeof token !== 'string') {
      return NextResponse.json({ error: 'token required' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 본인 user_id의 다른 토큰 모두 삭제 (현재 token은 유지)
    const { count, error } = await db.from('push_subscriptions')
      .delete({ count: 'exact' })
      .eq('user_id', user.id)
      .neq('token', token);

    if (error) {
      return NextResponse.json({ error: error.message }, { status: 500 });
    }

    return NextResponse.json({ ok: true, deleted: count ?? 0 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
