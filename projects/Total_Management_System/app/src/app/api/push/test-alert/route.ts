import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { sendOwnerAlert } from '@/lib/notification/owner-alert';

/**
 * POST /api/push/test-alert — 예비 알림 메일 테스트 (관리자 전용, 2026-09-29)
 * 설정 → 알림 → 「알림 메일 테스트」. 구글 재연결 직후 메일이 실제로 도착·울리는지 확인용.
 */
export async function POST() {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const r = await sendOwnerAlert(
    '[MAMORU] 테스트 · 앱알림 예비 메일',
    '앱 알림이 휴대폰에 도착하지 않을 때 이 형식으로 메일이 옵니다.\nGmail 앱 알림이 울렸다면 설정 완료입니다.',
    '/settings',
  );
  if (r.ok) {
    // 테스트 성공 = 설정이 고쳐졌다 → 이전 발송 실패 경고를 설정 화면에서 내린다 (api/push/devices)
    const { createServiceClient } = await import('@/lib/supabase/server');
    const now = new Date().toISOString();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await (createServiceClient() as any).from('system_settings')
      .upsert({ key: 'push.alert_test_ok_at', value: now, updated_at: now }, { onConflict: 'key' });
  }
  return NextResponse.json({ ok: r.ok, result: r.result });
}
