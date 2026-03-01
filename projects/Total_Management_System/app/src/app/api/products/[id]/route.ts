import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/products/[id] — 제품 상세 */
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

    const { data: product, error } = await db
      .from('products')
      .select('*')
      .eq('id', id)
      .single();

    if (error) throw error;

    // 매입처 정보
    let supplier = null;
    if (product.supplier_id) {
      const { data: sup } = await db
        .from('customers')
        .select('id, name')
        .eq('id', product.supplier_id)
        .single();
      supplier = sup;
    }

    return NextResponse.json({ product, supplier });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/products/[id] — 제품 수정 */
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
    const updates: Record<string, unknown> = {};

    const allowed = ['name', 'category', 'price', 'price_dealer', 'price_purchase', 'supplier_id', 'description', 'imweb_product_no', 'barcode', 'image_url', 'is_active'];
    for (const key of allowed) {
      if (key in body) updates[key] = body[key];
    }

    if (Object.keys(updates).length === 0) {
      return NextResponse.json({ error: '수정할 항목이 없습니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: product, error } = await db
      .from('products')
      .update(updates)
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;

    return NextResponse.json({ product });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
