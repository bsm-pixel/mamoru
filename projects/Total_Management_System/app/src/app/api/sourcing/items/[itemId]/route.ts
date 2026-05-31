import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/sourcing/items/[itemId] — 단건 품목 (입고매칭 페이지용) */
export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ itemId: string }> },
) {
  try {
    const { itemId } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data, error } = await db.from('sourcing_items').select('*').eq('id', itemId).single();
    if (error) throw error;

    return NextResponse.json({ item: data });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/**
 * PATCH /api/sourcing/items/[itemId] — 품목 수정
 * 입력/선별 모두 처리: product_name·unit_price·features_memo·vendor_url·moq,
 * inbound_photos·inbound_memo, inspection_status(pending|matched|selected|rejected)
 * inspection_status='selected' 로 바뀌면 selected_at 자동 스탬프.
 */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ itemId: string }> },
) {
  try {
    const { itemId } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const patch: Record<string, unknown> = { updated_at: new Date().toISOString() };
    for (const k of [
      'supplier_name', 'supplier_url', 'product_name', 'unit_price', 'features_memo', 'vendor_url', 'moq',
      'inbound_photos', 'inbound_memo', 'inspection_status',
    ]) {
      if (k in body) patch[k] = body[k];
    }
    if (body.inspection_status === 'selected') patch.selected_at = new Date().toISOString();
    if (body.inspection_status && body.inspection_status !== 'selected') patch.selected_at = null;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data, error } = await db
      .from('sourcing_items')
      .update(patch)
      .eq('id', itemId)
      .select()
      .single();
    if (error) throw error;

    return NextResponse.json({ item: data });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/sourcing/items/[itemId] — 품목 삭제 */
export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ itemId: string }> },
) {
  try {
    const { itemId } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { error } = await db.from('sourcing_items').delete().eq('id', itemId);
    if (error) throw error;

    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
