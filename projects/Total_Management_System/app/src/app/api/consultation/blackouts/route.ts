/**
 * /api/consultation/blackouts
 * 사장님 측 휴무일(closed_dates) 관리 API.
 *
 * 정책:
 *   - 관리자 전용 (auth 필수)
 *   - GET ?from=YYYY-MM-DD&to=YYYY-MM-DD → 그 기간의 휴무일 + 그 기간 confirmed/suggested 상담 카운트
 *   - POST { date, reason? } → closed_dates UPSERT
 *   - DELETE ?date=YYYY-MM-DD → closed_dates DELETE
 *
 * 사장님 룰: 막힘은 고객 셀프 예약 흐름에만 적용.
 *   사장님 측 흐름(admin-create, suggest, 일정수동등록 모달)은 closed_dates 무시.
 *   (memory/feedback_consultation_blackout_rule.md)
 */

import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

export async function GET(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const url = new URL(req.url);
  const from = url.searchParams.get('from');
  const to = url.searchParams.get('to');
  if (!from || !to) {
    return NextResponse.json({ error: 'from/to (YYYY-MM-DD) 필수' }, { status: 400 });
  }

  const db = createServiceClient();

  // 1) 휴무일 list
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data: blackouts, error: e1 } = await (db as any)
    .from('closed_dates')
    .select('date, reason')
    .gte('date', from)
    .lte('date', to)
    .order('date', { ascending: true });
  if (e1) return NextResponse.json({ error: e1.message }, { status: 500 });

  // 2) 그 기간 상담 일정 (충돌 표시용 — 매장/출장만, 톡상담 제외)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data: consultations, error: e2 } = await (db as any)
    .from('consultations')
    .select('id, visit_date, visit_time, consultation_type, status, name, phone')
    .gte('visit_date', from)
    .lte('visit_date', to)
    .in('status', ['confirmed', 'suggested', 'pending_admin', 'in_progress'])
    .in('consultation_type', ['store_visit', 'field_request'])
    .order('visit_date', { ascending: true });
  if (e2) return NextResponse.json({ error: e2.message }, { status: 500 });

  return NextResponse.json({
    blackouts: blackouts || [],
    consultations: consultations || [],
  });
}

export async function POST(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const body = await req.json().catch(() => null);
  const date = body?.date as string | undefined;
  const reason = (body?.reason as string | undefined) || null;
  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return NextResponse.json({ error: 'date (YYYY-MM-DD) 필수' }, { status: 400 });
  }

  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { error } = await (db as any)
    .from('closed_dates')
    .upsert({ date, reason }, { onConflict: 'date' });
  if (error) return NextResponse.json({ error: error.message }, { status: 500 });

  return NextResponse.json({ ok: true, date, reason });
}

export async function DELETE(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const url = new URL(req.url);
  const date = url.searchParams.get('date');
  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return NextResponse.json({ error: 'date (YYYY-MM-DD) 필수' }, { status: 400 });
  }

  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { error } = await (db as any).from('closed_dates').delete().eq('date', date);
  if (error) return NextResponse.json({ error: error.message }, { status: 500 });

  return NextResponse.json({ ok: true, date });
}
