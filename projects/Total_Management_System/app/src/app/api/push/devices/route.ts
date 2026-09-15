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

    const { data, error } = await db
      .from('push_subscriptions')
      .select('id, device_id, device_info, created_at, updated_at')
      .eq('user_id', user.id)
      .order('updated_at', { ascending: false });

    if (error) return NextResponse.json({ error: error.message }, { status: 500 });

    const devices = (data || []).map((d: Record<string, unknown>) => ({
      id: d.id,
      label: (d.device_info as string) || '이름 없는 기기 (구버전 등록)',
      isCurrent: !!currentDeviceId && d.device_id === currentDeviceId,
      updatedAt: d.updated_at,
    }));

    return NextResponse.json({ devices });
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
