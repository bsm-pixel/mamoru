import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

/** GET /api/events?status=received|payment_noticed|converted|cancelled|all */
export async function GET(req: NextRequest) {
  try {
    const status = req.nextUrl.searchParams.get('status') || 'all';
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let q = (db as any)
      .from('event_submissions')
      .select('*')
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
