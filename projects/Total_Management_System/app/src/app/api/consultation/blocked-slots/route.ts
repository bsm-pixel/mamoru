/**
 * /api/consultation/blocked-slots
 * 사장님 측 날짜별 시간대 차단(blocked_time_slots) 관리 API. (096)
 *
 * 정책:
 *   - 관리자 전용 (auth 필수)
 *   - GET ?from=YYYY-MM-DD&to=YYYY-MM-DD → 그 기간 날짜차단 + 매주 반복차단(요일) 전부
 *   - POST { date|weekday, start_time, end_time, reason? } → INSERT (단건 추가)
 *   - PUT  { date|weekday, ranges:[{start_time,end_time}] } → 그 날짜/요일 차단을 통째로 교체
 *          (어드민 시간 격자 저장용 — 칠한 칸을 구간으로 병합해 한 번에 반영, 118)
 *   - DELETE ?id=UUID → 단건 삭제 (한 날짜 여러 시간대 중 하나만)
 *
 * 사장님 룰: 시간대 차단은 고객 셀프 예약 흐름(매장/출장/톡 + 복원수리 직접방문)에만 적용.
 *   사장님 측 흐름(admin-create, suggest, 일정수동등록 모달)은 무시 — 항상 유동.
 *   (memory/feedback_consultation_blackout_rule.md)
 */

import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

const TIME_RE = /^([01]\d|2[0-3]):(00|30)$/; // 30분 단위 HH:MM

/** 'HH:MM' → 분. 겹침 검사용 */
function toMin(t: string): number {
  const [h, m] = t.split(':').map(Number);
  return h * 60 + m;
}
/** 반개구간 [a0,a1) 과 [b0,b1) 이 겹치나 */
function overlaps(a0: number, a1: number, b0: number, b1: number): boolean {
  return a0 < b1 && a1 > b0;
}

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
  const dbAny = db as any;
  // 날짜차단(기간) + 매주 반복차단(전체) 을 함께 반환 — 어드민 화면이 둘 다 그린다 (118)
  const [dateRes, weeklyRes] = await Promise.all([
    dbAny.from('blocked_time_slots')
      .select('id, date, weekday, start_time, end_time, reason')
      .gte('date', from).lte('date', to)
      .order('date', { ascending: true }).order('start_time', { ascending: true }),
    dbAny.from('blocked_time_slots')
      .select('id, date, weekday, start_time, end_time, reason')
      .not('weekday', 'is', null)
      .order('weekday', { ascending: true }).order('start_time', { ascending: true }),
  ]);
  if (dateRes.error) return NextResponse.json({ error: dateRes.error.message }, { status: 500 });
  if (weeklyRes.error) return NextResponse.json({ error: weeklyRes.error.message }, { status: 500 });

  return NextResponse.json({ blockedSlots: [...(dateRes.data || []), ...(weeklyRes.data || [])] });
}

export async function POST(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const body = await req.json().catch(() => null);
  const date = body?.date as string | undefined;
  const weekday = body?.weekday;
  const startTime = body?.start_time as string | undefined;
  const endTime = body?.end_time as string | undefined;
  const reason = (body?.reason as string | undefined) || null;

  const hasWeekday = typeof weekday === 'number' && Number.isInteger(weekday) && weekday >= 0 && weekday <= 6;
  const hasDate = !!date && /^\d{4}-\d{2}-\d{2}$/.test(date);
  if (hasDate === hasWeekday) {
    return NextResponse.json({ error: 'date(YYYY-MM-DD) 또는 weekday(0~6) 중 하나만 지정하세요' }, { status: 400 });
  }
  if (!startTime || !TIME_RE.test(startTime) || !endTime || !TIME_RE.test(endTime)) {
    return NextResponse.json({ error: 'start_time/end_time 은 30분 단위 HH:MM' }, { status: 400 });
  }
  if (startTime >= endTime) {
    return NextResponse.json({ error: '종료 시간이 시작 시간보다 늦어야 합니다' }, { status: 400 });
  }

  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  // 같은 날짜/요일에 이미 겹치는 구간이 있으면 거절 (중복·겹침 등록 방지)
  const dupQ = dbAny.from('blocked_time_slots').select('id, start_time, end_time');
  const { data: exist } = hasDate ? await dupQ.eq('date', date) : await dupQ.eq('weekday', weekday);
  const s = toMin(startTime), e = toMin(endTime);
  const clash = (exist || []).find((r: { start_time: string; end_time: string }) =>
    overlaps(s, e, toMin(r.start_time), toMin(r.end_time)));
  if (clash) {
    return NextResponse.json({ error: `이미 차단된 시간과 겹칩니다 (${clash.start_time}~${clash.end_time})` }, { status: 409 });
  }

  const { data, error } = await dbAny
    .from('blocked_time_slots')
    .insert({
      date: hasDate ? date : null,
      weekday: hasWeekday ? weekday : null,
      start_time: startTime, end_time: endTime, reason,
    })
    .select('id, date, weekday, start_time, end_time, reason')
    .single();
  if (error) return NextResponse.json({ error: error.message }, { status: 500 });

  return NextResponse.json({ ok: true, blockedSlot: data });
}

/**
 * PUT — 한 날짜(또는 한 요일)의 차단을 통째로 교체.
 * 어드민 시간 격자에서 칠한 칸을 구간으로 병합해 보내면, 기존 것을 지우고 새로 심는다.
 * body: { date?: 'YYYY-MM-DD', weekday?: 0..6, ranges: [{ start_time, end_time, reason? }] }
 */
export async function PUT(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const body = await req.json().catch(() => null);
  const date = body?.date as string | undefined;
  const weekday = body?.weekday;
  const ranges = Array.isArray(body?.ranges) ? body.ranges : null;

  const hasWeekday = typeof weekday === 'number' && Number.isInteger(weekday) && weekday >= 0 && weekday <= 6;
  const hasDate = !!date && /^\d{4}-\d{2}-\d{2}$/.test(date);
  if (hasDate === hasWeekday) {
    return NextResponse.json({ error: 'date 또는 weekday 중 하나만 지정하세요' }, { status: 400 });
  }
  if (!ranges) return NextResponse.json({ error: 'ranges 배열 필수' }, { status: 400 });

  // 검증 + 정렬 + 겹침 확인
  const rows: { start_time: string; end_time: string; reason: string | null }[] = [];
  for (const r of ranges) {
    const st = r?.start_time, et = r?.end_time;
    if (!st || !TIME_RE.test(st) || !et || !TIME_RE.test(et) || st >= et) {
      return NextResponse.json({ error: `잘못된 구간: ${st}~${et}` }, { status: 400 });
    }
    rows.push({ start_time: st, end_time: et, reason: (r?.reason as string) || null });
  }
  rows.sort((a, b) => a.start_time.localeCompare(b.start_time));
  for (let i = 1; i < rows.length; i++) {
    if (overlaps(toMin(rows[i - 1].start_time), toMin(rows[i - 1].end_time), toMin(rows[i].start_time), toMin(rows[i].end_time))) {
      return NextResponse.json({ error: '구간끼리 겹칩니다' }, { status: 400 });
    }
  }

  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  // 기존 것 제거 후 새로 삽입 (그 날짜/요일에 한해서만)
  const delQ = dbAny.from('blocked_time_slots').delete();
  const { error: delErr } = hasDate ? await delQ.eq('date', date) : await delQ.eq('weekday', weekday);
  if (delErr) return NextResponse.json({ error: delErr.message }, { status: 500 });

  if (rows.length > 0) {
    const payload = rows.map((r) => ({
      date: hasDate ? date : null,
      weekday: hasWeekday ? weekday : null,
      start_time: r.start_time, end_time: r.end_time, reason: r.reason,
    }));
    const { error: insErr } = await dbAny.from('blocked_time_slots').insert(payload);
    if (insErr) return NextResponse.json({ error: insErr.message }, { status: 500 });
  }

  return NextResponse.json({ ok: true, count: rows.length });
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
