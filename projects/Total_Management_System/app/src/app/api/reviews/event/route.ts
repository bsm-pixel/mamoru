import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

/**
 * 리뷰 이벤트 관리 (인증 필요, TMS 관리자용).
 *  GET  /api/reviews/event?month=YYMM → { config, reviews }  (그 달 설정 + 그 달 후기(당첨마킹 포함))
 *  POST /api/reviews/event            → { month, config, winners } 저장(설정 upsert + 당첨 마킹 재설정)
 */

/** 'YYMM' 의 KST 월 경계를 UTC ISO 로 (created_at timestamptz 필터용, toISOString UTC밀림 회피) */
function kstMonthRange(month: string): { startISO: string; endISO: string } | null {
  if (!/^\d{4}$/.test(month)) return null;
  const year = 2000 + parseInt(month.slice(0, 2), 10);
  const mon = parseInt(month.slice(2, 4), 10); // 1~12
  if (mon < 1 || mon > 12) return null;
  const KST = 9 * 3600 * 1000;
  const start = Date.UTC(year, mon - 1, 1) - KST;   // KST 그 달 1일 00:00 → UTC
  const end = Date.UTC(year, mon, 1) - KST;          // KST 다음 달 1일 00:00 → UTC
  return { startISO: new Date(start).toISOString(), endISO: new Date(end).toISOString() };
}

export async function GET(req: NextRequest) {
  try {
    const auth = await createServerSupabaseClient();
    const { data: { user } } = await auth.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    // 인증 확인 후 service role 로 DB 작업(review_event_config 는 RLS 켜짐 — authenticated 롤은 차단됨)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;

    const month = req.nextUrl.searchParams.get('month') || '';
    const range = kstMonthRange(month);
    if (!range) return NextResponse.json({ error: 'invalid month (YYMM)' }, { status: 400 });

    // 설정
    const { data: config } = await db
      .from('review_event_config')
      .select('*')
      .eq('month', month)
      .maybeSingle();

    // 그 달 후기(숨김 제외) — 당첨마킹은 event_month=this month 인 것만 유효로 노출
    const { data: reviews, error } = await db
      .from('reviews')
      .select('id, review_id, type, subtype, name, phone, stars, content, photo_urls, product, created_at, status, event_month, event_rank, event_display_name, event_route')
      .gte('created_at', range.startISO)
      .lt('created_at', range.endISO)
      .neq('status', 'hidden')
      .order('created_at', { ascending: false });
    if (error) throw error;

    return NextResponse.json({ config: config || null, reviews: reviews || [] });
  } catch (err) {
    console.error('[reviews/event] GET 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

interface WinnerInput { id: string; rank: number; display_name?: string | null; route?: string | null }

export async function POST(req: NextRequest) {
  try {
    const auth = await createServerSupabaseClient();
    const { data: { user } } = await auth.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    // 인증 확인 후 service role 로 DB 작업(RLS 우회 — 관리자 엔드포인트)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;

    const body = await req.json();
    const month: string = body.month;
    if (!/^\d{4}$/.test(month || '')) return NextResponse.json({ error: 'invalid month (YYMM)' }, { status: 400 });
    const config = body.config || {};
    const winners: WinnerInput[] = Array.isArray(body.winners) ? body.winners : [];

    // 1) 설정 upsert
    const { error: cfgErr } = await db
      .from('review_event_config')
      .upsert({
        month,
        deadline: config.deadline || null,
        announce_at: config.announce_at || null,
        hero_image_url: config.hero_image_url || null,
        prizes: Array.isArray(config.prizes) ? config.prizes : [],
        status: config.status || 'draft',
        updated_at: new Date().toISOString(),
      }, { onConflict: 'month' });
    if (cfgErr) throw cfgErr;

    // 2) 당첨 마킹 재설정 — 이 달 기존 마킹 전부 해제 후 선택분만 다시 설정(멱등)
    const { error: clearErr } = await db
      .from('reviews')
      .update({ event_month: null, event_rank: null, event_display_name: null, event_route: null })
      .eq('event_month', month);
    if (clearErr) throw clearErr;

    for (const w of winners) {
      if (!w.id || !w.rank) continue;
      const { error: upErr } = await db
        .from('reviews')
        .update({
          event_month: month,
          event_rank: w.rank,
          event_display_name: (w.display_name || '').trim() || null,
          event_route: (w.route || '').trim() || null,
        })
        .eq('id', w.id);
      if (upErr) throw upErr;
    }

    return NextResponse.json({ ok: true, saved: winners.length });
  } catch (err) {
    console.error('[reviews/event] POST 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
