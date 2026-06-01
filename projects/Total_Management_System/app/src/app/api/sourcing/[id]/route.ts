import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/sourcing/[id] — 발주 + 품목 (정렬순) */
export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const [poRes, itemsRes] = await Promise.all([
      db.from('sourcing_pos').select('*').eq('id', id).single(),
      db.from('sourcing_items').select('*').eq('po_id', id).order('sort_order', { ascending: true }),
    ]);
    if (poRes.error) throw poRes.error;

    // 연결된 제품/부자재 머지 (뱃지용 — 임베드 대신 2-query로 FK 타이밍 무관 안전)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const items: any[] = itemsRes.data || [];
    const linkedIds = [...new Set(items.map((it) => it.linked_product_id).filter(Boolean))];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const prodMap: Record<string, any> = {};
    if (linkedIds.length > 0) {
      const { data: prods } = await db.from('products').select('id, sku, name, category').in('id', linkedIds);
      for (const p of prods || []) prodMap[p.id] = p;
    }
    const itemsWithLinked = items.map((it) => ({
      ...it,
      linked_product: it.linked_product_id ? prodMap[it.linked_product_id] ?? null : null,
    }));

    return NextResponse.json({ po: poRes.data, items: itemsWithLinked });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/sourcing/[id] — 발주 헤더 수정 */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const patch: Record<string, unknown> = { updated_at: new Date().toISOString() };
    for (const k of ['supplier_name', 'supplier_url', 'order_date', 'exchange_rate', 'status', 'memo']) {
      if (k in body) patch[k] = body[k];
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { error } = await db.from('sourcing_pos').update(patch).eq('id', id);
    if (error) throw error;

    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/sourcing/[id] — 발주 삭제 (품목 CASCADE) */
export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { error } = await db.from('sourcing_pos').delete().eq('id', id);
    if (error) throw error;

    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
