import { NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { syncConsultationToCalendar } from '@/lib/google/calendar-sync';

/**
 * GET /api/consultation/admin-cleanup-calendar
 *
 * cancelled 상태인데 google_event_id가 아직 남아있는 잔여 건들을
 * 일괄로 syncConsultationToCalendar 호출하여 Google Calendar에서 삭제 + DB 정리.
 *
 * cancel API에 캘린더 동기화가 누락됐던 시기에 누적된 잔여 데이터를 일회성 정리.
 * 향후 cancel API가 자동 동기화하므로 보통 결과가 0건이어야 정상.
 *
 * syncConsultationToCalendar는 idempotent — 이미 정리된 건은 무시.
 */
export async function GET() {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  const { data: dirty, error } = await dbAny
    .from('consultations')
    .select('id, name, consultation_type, status, google_event_id, visit_date, visit_time')
    .eq('status', 'cancelled')
    .not('google_event_id', 'is', null)
    .order('visit_date', { ascending: false });

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  const items = (dirty || []) as Array<{
    id: string;
    name: string;
    consultation_type: string;
    visit_date: string | null;
    visit_time: string | null;
  }>;

  const results: Array<{
    id: string;
    name: string;
    type: string;
    visit_date: string | null;
    visit_time: string | null;
    success: boolean;
    error?: string;
  }> = [];

  for (const item of items) {
    try {
      await syncConsultationToCalendar(item.id);
      results.push({
        id: item.id,
        name: item.name,
        type: item.consultation_type,
        visit_date: item.visit_date,
        visit_time: item.visit_time,
        success: true,
      });
    } catch (err) {
      results.push({
        id: item.id,
        name: item.name,
        type: item.consultation_type,
        visit_date: item.visit_date,
        visit_time: item.visit_time,
        success: false,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  return NextResponse.json({
    ok: true,
    processed: results.length,
    succeeded: results.filter((r) => r.success).length,
    failed: results.filter((r) => !r.success).length,
    items: results,
  });
}
