import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { isMobileLabel } from '@/lib/firebase/push-fallback';

/**
 * POST /api/push/ack — 서비스워커의 "알림 받았음" 신호 (2026-09-29, 마이그 155)
 *
 * 🔓 로그인 불필요: 로그인이 풀린 휴대폰도 알림은 받을 수 있고, 그 수신도 인정돼야 한다.
 *    받는 값은 알림 id(nid)·기기 id 뿐이고 하는 일은 "수신 시각 기록"이 전부라 노출 위험이 없다.
 *
 * 휴대폰 수신만 push_notifications.acked_at 으로 인정한다 → PC만 받고 폰이 못 받으면 메일이 간다.
 * body: { nid, deviceId?, ua? }
 */
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export async function POST(req: NextRequest) {
  try {
    const { nid, deviceId, ua } = await req.json().catch(() => ({}));
    if (typeof nid !== 'string' || !UUID.test(nid)) {
      return NextResponse.json({ error: 'nid required' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const now = new Date().toISOString();

    // 기기 이름: 등록된 기기 행 우선, 없으면(기기 id 못 읽음) SW 의 userAgent 로 판별
    let label = '';
    if (typeof deviceId === 'string' && deviceId) {
      const { data } = await db.from('push_subscriptions')
        .update({ last_ack_at: now })
        .eq('device_id', deviceId)
        .select('device_info');
      label = data?.[0]?.device_info || '';
    }
    if (!label && typeof ua === 'string') {
      label = /Android/i.test(ua) ? 'Android' : /iPhone/i.test(ua) ? 'iPhone' : /iPad/i.test(ua) ? 'iPad' : 'PC';
    }

    if (isMobileLabel(label)) {
      await db.from('push_notifications')
        .update({ acked_at: now, acked_device: label })
        .eq('id', nid)
        .is('acked_at', null);
    }

    return NextResponse.json({ ok: true, mobile: isMobileLabel(label) });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
