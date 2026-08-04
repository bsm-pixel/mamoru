import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { syncRepairToCalendar } from '@/lib/google/repair-calendar-sync';

/**
 * GET /api/cron/repair-calendar-backfill
 *
 * 직접방문인데 구글 캘린더 미기록(google_event_id NULL)인 '예정' 건을 주기적으로 자동 재동기화(self-heal).
 * - 근본수정(접수/PATCH after→await, 2026-08-04) 이후에도 만일의 유실을 크론이 안전망으로 복구.
 * - 대상: proceed_type='직접방문' AND google_event_id IS NULL AND status NOT IN(취소/완료/배송완료) AND 방문일 >= 오늘(KST)
 * - 구글 인증은 Vercel 서버의 기존 연결(system_settings 토큰)을 사용 — 별도 키 불필요.
 *   연결이 끊긴 경우(refresh_token 폐기 등)엔 sync 가 조용히 skip → TMS 설정에서 구글 재연결 필요.
 */
export async function GET(request: NextRequest) {
  const authHeader = request.headers.get('authorization');
  if (authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = createServiceClient() as any;
  const todayKST = new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 10); // 'YYYY-MM-DD'(KST)

  const { data, error } = await db
    .from('repairs')
    .select('id, as_id, visit_date, visit_time, status, google_event_id')
    .eq('proceed_type', '직접방문')
    .is('google_event_id', null)
    .not('status', 'in', '("cancelled","completed","delivered")')
    .gte('visit_date', todayKST);
  if (error) return NextResponse.json({ error: String(error) }, { status: 500 });

  const targets: Array<{ id: string; as_id: string }> = data || [];
  const results: Array<{ as_id: string; synced: boolean }> = [];

  for (const r of targets) {
    try {
      await syncRepairToCalendar(r.id); // 내부에서 예외 삼킴 — 여기선 결과만 확인
      const { data: after } = await db.from('repairs').select('google_event_id').eq('id', r.id).single();
      results.push({ as_id: r.as_id, synced: !!after?.google_event_id });
    } catch {
      results.push({ as_id: r.as_id, synced: false });
    }
  }

  const healed = results.filter((x) => x.synced).length;
  if (targets.length > 0) console.log(`[repair-calendar-backfill] 점검 ${targets.length}건 / 복구 ${healed}건`, results);
  return NextResponse.json({ checked: targets.length, healed, results });
}
