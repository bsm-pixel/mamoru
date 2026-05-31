import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/sourcing/[id]/items — 발주에 품목 1건 추가 (sticker_no 자동 채번) */
export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: po, error: poErr } = await db
      .from('sourcing_pos')
      .select('po_number')
      .eq('id', id)
      .single();
    if (poErr) throw poErr;

    // 기존 품목 수 → 다음 순번
    const { data: existing } = await db
      .from('sourcing_items')
      .select('id')
      .eq('po_id', id);
    const next = (existing?.length || 0) + 1;
    const sticker_no = `${po.po_number}-${String(next).padStart(3, '0')}`;

    const { data: item, error } = await db
      .from('sourcing_items')
      .insert({
        po_id: id,
        sticker_no,
        supplier_name: body.supplier_name || null,
        supplier_url: body.supplier_url || null,
        vendor_url: body.vendor_url || null,
        product_name: body.product_name || '',
        features_memo: body.features_memo || null,
        unit_price: body.unit_price ?? 0,
        moq: body.moq ?? null,
        sort_order: next - 1,
      })
      .select()
      .single();
    if (error) throw error;

    return NextResponse.json({ item }, { status: 201 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
