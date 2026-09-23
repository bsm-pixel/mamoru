import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { syncConsultationToCalendar, loadFormatSettings } from '@/lib/google/calendar-sync';
import { syncRepairToCalendar } from '@/lib/google/repair-calendar-sync';
import { formatConsultationToEvent, formatRepairToEvent, formatShippingTodoToTask } from '@/lib/google/event-formatter';

const BASE_URL = process.env.NEXT_PUBLIC_APP_URL || 'https://app-eta-sandy-75.vercel.app';

const CRON_SECRET = process.env.CRON_SECRET || 'mamoru-tms-cron-2026';

/**
 * GET /api/cron/calendar-reformat — 이미 등록된 캘린더 일정을 현재 포맷으로 다시 쓰기 (2026-09-24)
 *
 * 포맷(제목·설명)을 바꿔도 **옛 일정은 그대로 남는다** — 동기화는 상태가 바뀔 때만 돌기 때문.
 * 이 경로가 '앞으로 남은 일정'만 골라 한 번 더 PATCH 해서 형식을 맞춘다.
 *   · 상담: google_event_id 있음 + 방문일 >= 오늘(KST) + 취소 아님
 *   · 복원수리 직접방문: 같은 조건
 * sync 함수는 idempotent(있으면 PATCH, 사라졌으면 재생성) — 여러 번 돌려도 안전.
 *
 * 크론에 넣지 않았다 — 포맷을 바꾼 직후 수동으로 한 번만 부르면 되는 정리용.
 * ?dry=1 이면 대상만 센다.
 */
export async function GET(req: NextRequest) {
  if (req.headers.get('authorization') !== `Bearer ${CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }
  const dry = req.nextUrl.searchParams.get('dry') === '1';
  const todayKST = new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 10);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = createServiceClient() as any;

  const { data: consults } = await db
    .from('consultations')
    .select('id, name, consultation_type, visit_date, status')
    .not('google_event_id', 'is', null)
    .neq('status', 'cancelled')
    .gte('visit_date', todayKST);

  const { data: repairs } = await db
    .from('repairs')
    .select('id, as_id, visit_date, status')
    .eq('proceed_type', '직접방문')
    .not('google_event_id', 'is', null)
    .not('status', 'in', '("cancelled")')
    .gte('visit_date', todayKST);

  const targets = {
    consultations: (consults || []).map((c: { id: string; visit_date: string }) => `${c.visit_date} ${c.id.slice(0, 8)}`),
    repairs: (repairs || []).map((r: { as_id: string; visit_date: string }) => `${r.visit_date} ${r.as_id}`),
  };
  if (dry) return NextResponse.json({ ok: true, dry: true, ...targets });

  // ?preview=1 — 구글에 쓰지 않고 '저장될 제목·설명'만 그대로 보여준다 (포맷 확인용)
  if (req.nextUrl.searchParams.get('preview') === '1') {
    const settings = await loadFormatSettings();
    const show = (e: { summary?: string | null; description?: string | null; location?: string | null }) =>
      ({ summary: e.summary, location: e.location || null, description: e.description });
    const { data: cRows } = await db.from('consultations').select('*')
      .not('google_event_id', 'is', null).neq('status', 'cancelled').gte('visit_date', todayKST).limit(5);
    let { data: rRows } = await db.from('repairs').select('*')
      .eq('proceed_type', '직접방문').not('google_event_id', 'is', null).gte('visit_date', todayKST).limit(3);
    // 예정된 직접방문이 없을 때도 형식은 확인할 수 있게 — 가장 최근 건으로 대체
    if (!rRows || rRows.length === 0) {
      const fallback = await db.from('repairs').select('*')
        .eq('proceed_type', '직접방문').not('visit_date', 'is', null)
        .order('visit_date', { ascending: false }).limit(1);
      rRows = fallback.data || [];
    }
    const { data: dRows } = await db.from('deliveries').select('*')
      .eq('status', 'confirmed').is('tracking_number', null).is('cancelled_at', null).limit(2);
    return NextResponse.json({
      ok: true,
      preview: true,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      consultations: (cRows || []).map((c: any) => show(formatConsultationToEvent({ ...c, adminNote: c.admin_note }, settings, BASE_URL))),
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      repairs: (rRows || []).map((r: any) => show(formatRepairToEvent(r, settings, BASE_URL))),
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      // 송장 대기건은 '일정'이 아니라 구글 할 일(Tasks) 로 나간다 — 제목·메모만 확인
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      shippingTodos: (dRows || []).map((d: any) => formatShippingTodoToTask({
        kind: 'delivery', who: d.customer_name || '거래처', docNo: d.dl_number, amount: d.total_amount,
        createdAt: d.created_at,
      })),
    });
  }

  for (const c of consults || []) await syncConsultationToCalendar(c.id);
  for (const r of repairs || []) await syncRepairToCalendar(r.id);

  return NextResponse.json({ ok: true, rewritten: targets });
}
