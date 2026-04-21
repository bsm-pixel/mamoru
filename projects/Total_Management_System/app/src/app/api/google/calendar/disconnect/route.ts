/**
 * POST /api/google/calendar/disconnect
 * Google Calendar 연결 해제 — 토큰 삭제 (이미 생성된 이벤트는 보존)
 */

import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { disconnectCalendar } from '@/lib/google/oauth';

export async function POST() {
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  try {
    await disconnectCalendar();
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
