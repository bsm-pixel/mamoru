/**
 * /api/consultation/blocked-slots
 * 사장님 측 날짜별 시간대 차단(blocked_time_slots) 관리 API. (096)
 *
 * 정책:
 *   - 관리자 전용 (auth 필수)
 *   - GET ?from=YYYY-MM-DD&to=YYYY-MM-DD → 그 기간 시간대 차단 목록
 *   - POST { date, start_time, end_time, reason? } → blocked_time_slots INSERT
 *   - DELETE ?id=UUID → 단건 삭제 (한 날짜 여러 시간대 중 하나만)
 *
 * 사장님 룰: 시간대 차단은 고객 셀프 예약 흐름(매장/출장/톡 + 복원수리 직접방문)에만 적용.
 *   사장님 측 흐름(admin-create, suggest, 일정수동등록 모달)은 무시 — 항상 유동.
 *   (memory/feedback_consultation_blackout_rule.md)
 */

import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

const TIME_RE = /^([01]\d|2[0-3]):(00|30)$/; // 30분 단위 HH:MM

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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data, error } = await (db as any)
    .from('blocked_time_slots')
    .select('id, date, start_time, end_time, reason')
    .gte('date', from)
    .lte('date', to)
    .order('date', { ascending: true })
    .order('start_time', { ascending: true });
  if (error) return NextResponse.json({ error: error.message }, { status: 500 });

  return NextResponse.json({ blockedSlots: data || [] });
}

export async function POST(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const body = await req.json().catch(() => null);
  const date = body?.date as string | undefined;
  const startTime = body?.start_time as string | undefined;
  const endTime = body?.end_time as string | undefined;
  const reason = (body?.reason as string | undefined) || null;

  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return NextResponse.json({ error: 'date (YYYY-MM-DD) 필수' }, { status: 400 });
  }
  if (!startTime || !TIME_RE.test(startTime) || !endTime || !TIME_RE.test(endTime)) {
    return NextResponse.json({ error: 'start_time/end_time 은 30분 단위 HH:MM' }, { status: 400 });
  }
  // 종료가 시작보다 늦어야 함
  if (startTime >= endTime) {
    return NextResponse.json({ error: '종료 시간이 시작 시간보다 늦어야 합니다' }, { status: 400 });
  }

  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data, error } = await (db as any)
    .from('blocked_time_slots')
    .insert({ date, start_time: startTime, end_time: endTime, reason })
    .select('id, date, start_time, end_time, reason')
    .single();
  if (error) return NextResponse.json({ error: error.message }, { status: 500 });

  return NextResponse.json({ ok: true, blockedSlot: data });
}

export async function DELETE(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const url = new URL(req.url);
  const id = url.searchParams.get('id');
  if (!id) {
    return NextResponse.json({ error: 'id 필수' }, { status: 400 });
  }

  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { error } = await (db as any).from('blocked_time_slots').delete().eq('id', id);
  if (error) return NextResponse.json({ error: error.message }, { status: 500 });

  return NextResponse.json({ ok: true, id });
}
