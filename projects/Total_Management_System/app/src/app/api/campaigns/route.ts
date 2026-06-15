import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

/** GET /api/campaigns — 캠페인 목록 (CORS도 허용: 고객 폼에서 캠페인명 표시용) */
const CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' };
export function OPTIONS() { return new NextResponse(null, { status: 204, headers: CORS }); }

export async function GET() {
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (db as any)
      .from('event_campaigns')
      .select('*')
      .order('created_at', { ascending: false });
    if (error) throw error;
    return NextResponse.json({ ok: true, campaigns: data || [] }, { headers: CORS });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS });
  }
}

/** POST /api/campaigns — 캠페인 생성 */
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    if (!body.name?.trim()) return NextResponse.json({ ok: false, error: '캠페인명 필수' }, { status: 400 });
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (db as any)
      .from('event_campaigns')
      .insert({
        name: body.name.trim(),
        type: body.type || 'other',
        status: 'active',
        starts_at: body.starts_at || null,
        ends_at: body.ends_at || null,
        memo: body.memo || null,
      })
      .select()
      .single();
    if (error) throw error;
    return NextResponse.json({ ok: true, campaign: data });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
