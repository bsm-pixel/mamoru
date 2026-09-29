import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * 알림 받는 기기 목록 (2026-09-15 신규)
 *
 * 왜 만들었나: 전엔 "내 알림이 몇 대에 등록돼 있는지" 볼 방법이 아예 없었다.
 * 그래서 PC·모바일 중 한쪽만 오는데도 원인을 화면에서 확인할 수 없었다.
 *   GET    → 등록된 기기 목록
 *   DELETE → 기기 1개 해제 (?id=)
 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const currentDeviceId = req.nextUrl.searchParams.get('deviceId') || '';

    // 2026-09-29: last_ack_at(마지막 실제 수신) 추가 — 마이그 155 전이면 옛 컬럼만
    let { data, error } = await db
      .from('push_subscriptions')
      .select('id, device_id, device_info, created_at, updated_at, last_ack_at')
      .eq('user_id', user.id)
      .order('updated_at', { ascending: false });
    if (error) {
      ({ data, error } = await db
        .from('push_subscriptions')
        .select('id, device_id, device_info, created_at, updated_at')
        .eq('user_id', user.id)
        .order('updated_at', { ascending: false }));
    }

    if (error) return NextResponse.json({ error: error.message }, { status: 500 });

    const devices = (data || []).map((d: Record<string, unknown>) => ({
      id: d.id,
      label: (d.device_info as string) || '이름 없는 기기 (구버전 등록)',
      isCurrent: !!currentDeviceId && d.device_id === currentDeviceId,
      updatedAt: d.updated_at,
      lastAckAt: (d.last_ack_at as string) || null,
    }));

    // 예비 메일 안전망 상태 — 설정 화면에서 "안전망이 켜져 있는가"를 한눈에
    const { getOwnerAlertStatus } = await import('@/lib/notification/owner-alert');
    const { isSweepAlive } = await import('@/lib/firebase/push-fallback');
    const { createServiceClient } = await import('@/lib/supabase/server');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const svc = createServiceClient() as any;
    const since = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();
    const { count: fallbacks7d } = await svc.from('push_notifications')
      .select('id', { count: 'exact', head: true })
      .gte('created_at', since)
      .not('fallback_sent_at', 'is', null);
    const alert = await getOwnerAlertStatus();
    // 권한은 있어도 실제 발송이 실패했을 수 있다(Gmail API 꺼짐 등) → 마지막 예비발송 결과로 조용한 실패를 드러낸다
    const { data: lastFallback } = await svc.from('push_notifications')
      .select('fallback_result, fallback_sent_at')
      .not('fallback_sent_at', 'is', null)
      .order('fallback_sent_at', { ascending: false })
      .limit(1);
    const lastResult: string = lastFallback?.[0]?.fallback_result || '';
    const { data: testOk } = await svc.from('system_settings').select('value').eq('key', 'push.alert_test_ok_at').maybeSingle();
    const testOkAt = testOk?.value ? String(testOk.value).replace(/^"|"$/g, '') : '';
    const lastFailed = !!lastResult && lastResult !== 'mailed'
      && !(testOkAt && Date.parse(testOkAt) > Date.parse(String(lastFallback?.[0]?.fallback_sent_at || 0)));   // 이후 테스트 메일 성공이면 해소된 것
    const lastProblem = lastResult.startsWith('gmail-api-disabled')
      ? '구글 클라우드에서 Gmail API 가 꺼져 있습니다 — 콘솔에서 사용 설정 필요.'
      : `마지막 알림 메일 발송이 실패했습니다 (${lastResult.slice(0, 60)}).`;

    return NextResponse.json({
      devices,
      safety: {
        alertReady: alert.ready && !lastFailed,
        alertProblem: alert.problem || (lastFailed ? lastProblem : null),
        sweepAlive: await isSweepAlive(svc),
        fallbacks7d: fallbacks7d ?? null,
      },
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

export async function DELETE(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const id = req.nextUrl.searchParams.get('id');
    if (!id) return NextResponse.json({ error: 'id required' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    // 본인 소유 행만 삭제 (user_id 조건 필수)
    const { error } = await db.from('push_subscriptions').delete().eq('id', id).eq('user_id', user.id);
    if (error) return NextResponse.json({ error: error.message }, { status: 500 });

    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
