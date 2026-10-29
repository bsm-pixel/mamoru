import { NextRequest, NextResponse } from 'next/server';
import { requireUser } from '@/lib/auth/require-user';
import { createServiceClient } from '@/lib/supabase/server';

/** GET /api/events?status=…&kind=event|stock_sale
 *  kind 기본값 'event' — 재고판매(stock_sale)와 화면을 분리한다 (117) */
export async function GET(req: NextRequest) {
  const deny = await requireUser(); if (deny) return deny; // 2026-10-29 로그인 없이 열려 있던 관리자 API 차단
  try {
    const status = req.nextUrl.searchParams.get('status') || 'all';
    const kind = req.nextUrl.searchParams.get('kind') || 'event';
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let q = (db as any)
      .from('event_submissions')
      .select('*')
      .eq('kind', kind)
      .order('created_at', { ascending: false });
    if (status !== 'all') q = q.eq('status', status);
    const { data, error } = await q.limit(500);
    if (error) throw error;
    return NextResponse.json({ ok: true, events: data || [] });
  } catch (err) {
    console.error('[events GET] 실패:', err);
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
