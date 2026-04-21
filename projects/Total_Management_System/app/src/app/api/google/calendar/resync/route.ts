/**
 * POST /api/google/calendar/resync
 * 전체 재동기화 — 활성 상담(confirmed/reschedule_requested/change_requested/completed)을
 * Google Calendar에 일괄 생성/업데이트
 *
 * 제약:
 *   - 과거 60일 ~ 미래 180일 범위만
 *   - talk_consult 제외
 *   - 한 요청당 최대 200건
 */

import { NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { syncConsultationToCalendar } from '@/lib/google/calendar-sync';
import { getConnectionStatus } from '@/lib/google/oauth';

const SYNC_STATUSES = ['confirmed', 'reschedule_requested', 'change_requested', 'completed'];

export async function POST() {
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  // 연결 확인
  const status = await getConnectionStatus();
  if (!status.connected) {
    return NextResponse.json({ ok: false, error: 'Google Calendar가 연결되지 않았습니다' }, { status: 400 });
  }

  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 과거 60일 ~ 미래 180일
    const today = new Date();
    const minDate = new Date(today);
    minDate.setDate(minDate.getDate() - 60);
    const maxDate = new Date(today);
    maxDate.setDate(maxDate.getDate() + 180);

    const minStr = minDate.toISOString().slice(0, 10);
    const maxStr = maxDate.toISOString().slice(0, 10);

    const { data: rows, error } = await dbAny
      .from('consultations')
      .select('id')
      .in('status', SYNC_STATUSES)
      .neq('consultation_type', 'talk_consult')
      .gte('visit_date', minStr)
      .lte('visit_date', maxStr)
      .order('visit_date', { ascending: true })
      .limit(200);

    if (error) {
      return NextResponse.json({ ok: false, error: error.message }, { status: 500 });
    }

    const ids = (rows || []).map((r: { id: string }) => r.id);

    // 순차 처리 (Google API rate limit 고려)
    let success = 0;
    let failed = 0;
    for (const id of ids) {
      try {
        await syncConsultationToCalendar(id);
        success++;
      } catch {
        failed++;
      }
    }

    return NextResponse.json({
      ok: true,
      data: {
        total: ids.length,
        success,
        failed,
        range: { from: minStr, to: maxStr },
      },
    });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
