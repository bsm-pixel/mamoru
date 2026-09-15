import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * POST /api/push/subscribe — FCM 토큰 등록 (기기당 1개)
 *
 * 🔴 2026-09-15 수정 — PC·모바일이 동시에 알림을 못 받던 근본원인 제거.
 *    전에는 "사용자당 토큰 1개" 정책이라 등록할 때마다
 *      DELETE WHERE user_id = ? AND token <> ?
 *    를 돌려서 **다른 기기의 토큰을 지웠다.**
 *    → PC에서 열면 모바일이 끊기고, 모바일에서 열면 PC가 끊겼다(실측: 등록기기 1대).
 *
 *    원래 의도(같은 기기의 토큰 중복 누적 차단)는 기기 단위로 처리해야 맞다.
 *    이제 device_id(브라우저 localStorage UUID) 기준으로 그 **기기의 옛 토큰만** 교체한다.
 *
 * body: { token, deviceId?, deviceInfo? }
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { token, deviceId, deviceInfo } = await req.json();
    if (!token) return NextResponse.json({ error: 'token required' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const now = new Date().toISOString();

    // 마이그레이션 151 적용 전(배포가 먼저 올라간 경우)에도 구독이 깨지지 않게 감싼다.
    if (deviceId) {
      // 1) 같은 기기의 옛 토큰만 제거 — 다른 기기는 건드리지 않는다
      const { error: delErr } = await db.from('push_subscriptions')
        .delete()
        .eq('user_id', user.id)
        .eq('device_id', deviceId)
        .neq('token', token);

      if (!delErr) {
        const { error: upErr } = await db.from('push_subscriptions').upsert(
          {
            user_id: user.id,
            token,
            device_id: deviceId,
            device_info: deviceInfo || null,
            last_seen_at: now,
            updated_at: now,
          },
          { onConflict: 'token' },
        );
        if (!upErr) return NextResponse.json({ ok: true, mode: 'per-device' });
        console.warn('[push/subscribe] per-device upsert 실패 → 레거시 경로:', upErr.message);
      } else {
        console.warn('[push/subscribe] device_id 컬럼 없음(마이그 151 미적용?) → 레거시 경로:', delErr.message);
      }
    }

    // ── 레거시 폴백: device_id 없이 토큰만 저장 (다른 기기 토큰은 지우지 않는다) ──
    await db.from('push_subscriptions').upsert(
      { user_id: user.id, token, device_info: deviceInfo || null, updated_at: now },
      { onConflict: 'token' },
    );

    return NextResponse.json({ ok: true, mode: 'legacy' });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
