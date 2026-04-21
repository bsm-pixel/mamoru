/**
 * GET /api/google/calendar/status
 * Google Calendar 연결 상태 조회 (설정 UI용)
 */

import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getConnectionStatus } from '@/lib/google/oauth';

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
    return NextResponse.json({ ok: true, data: status });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
