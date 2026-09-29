import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { isMobileLabel } from '@/lib/firebase/push-fallback';
import { sendOwnerSms } from '@/lib/notification/owner-sms';

/**
 * POST /api/push/lost — 서비스워커가 "푸시 구독이 교체돼 이 기기 알림이 끊겼다"고 알림 (2026-09-29)
 *
 * SW 안에서는 새 FCM 토큰을 받을 수 없어 스스로 복구가 안 된다 → 사장님께 "TMS 앱 한 번 열기" 문자.
 * (앱을 열면 로그인 여부와 무관하게 토큰이 재등록된다 — 로그인 화면 포함)
 *
 * 🔓 무인증이라 남용 방지: 등록된 **휴대폰** 기기 id 일 때만, 그리고 6시간에 1번만 문자.
 * body: { deviceId }
 */
const THROTTLE_KEY = 'push.lost_warned_at';
const THROTTLE_MS = 6 * 60 * 60 * 1000;

export async function POST(req: NextRequest) {
  try {
    const { deviceId } = await req.json().catch(() => ({}));
    if (typeof deviceId !== 'string' || !deviceId) return NextResponse.json({ ok: false });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const { data: rows } = await db.from('push_subscriptions').select('device_info').eq('device_id', deviceId).limit(1);
    const label: string = rows?.[0]?.device_info || '';
    if (!isMobileLabel(label)) return NextResponse.json({ ok: false, reason: 'not-mobile-device' });

    const { data: last } = await db.from('system_settings').select('value').eq('key', THROTTLE_KEY).maybeSingle();
    const lastAt = last?.value ? new Date(String(last.value).replace(/^"|"$/g, '')).getTime() : 0;
    if (lastAt && Date.now() - lastAt < THROTTLE_MS) return NextResponse.json({ ok: true, throttled: true });

    const nowIso = new Date().toISOString();
    await db.from('system_settings').upsert({ key: THROTTLE_KEY, value: nowIso, updated_at: nowIso }, { onConflict: 'key' });
    const r = await sendOwnerSms(`[MAMORU] ${label} 앱알림 연결이 끊겼습니다.\nTMS 앱을 한 번 열어주세요(로그인 안 해도 됨).`);
    return NextResponse.json({ ok: r.ok });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
