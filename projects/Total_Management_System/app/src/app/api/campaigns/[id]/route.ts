import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

/** PATCH /api/campaigns/[id] — 캠페인 수정 (이름/유형/상태/할인규칙) */
export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const body = await req.json();
    const patch: Record<string, unknown> = {};
    if (typeof body.name === 'string') patch.name = body.name.trim();
    if (typeof body.type === 'string') patch.type = body.type;
    if (typeof body.status === 'string') patch.status = body.status;
    if (typeof body.memo === 'string') patch.memo = body.memo;
    if (Array.isArray(body.discount_rules)) patch.discount_rules = body.discount_rules;
    if (body.starts_at !== undefined) patch.starts_at = body.starts_at || null;
    if (body.ends_at !== undefined) patch.ends_at = body.ends_at || null;
    if (Object.keys(patch).length === 0) return NextResponse.json({ ok: true });

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { error } = await (db as any).from('event_campaigns').update(patch).eq('id', id);
    if (error) throw error;
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
