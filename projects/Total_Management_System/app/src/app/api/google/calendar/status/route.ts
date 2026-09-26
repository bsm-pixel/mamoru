/**
 * GET /api/google/calendar/status
 * Google Calendar 연결 상태 조회 (설정 UI용)
 */

import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getConnectionStatus } from '@/lib/google/oauth';
import { probeTasks } from '@/lib/google/tasks-client';

export async function GET() {
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  try {
    const status = await getConnectionStatus();

    /* 권한이 있어도 구글 클라우드에서 API 가 꺼져 있으면 동작하지 않는다(2026-09-26 실제로 걸림).
       '권한 있음 = 정상'으로 끝내면 또 조용히 실패하므로, 연결돼 있을 때 한 번 찔러보고 결과를 같이 내려준다. */
    let tasksApiDisabled = false;
    if (status.connected && !status.needs_reauth) {
      const probe = await probeTasks();
      tasksApiDisabled = !!probe.apiDisabled;
    }
    return NextResponse.json({ ok: true, data: { ...status, tasks_api_disabled: tasksApiDisabled } });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
