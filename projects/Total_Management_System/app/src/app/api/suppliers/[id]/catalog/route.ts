import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/suppliers/[id]/catalog — 매입품목 조회 */
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

    const { data, error } = await db
      .from('supplier_product_catalog')
      .select('id, supplier_id, product_id, order_name, features, created_at, products:product_id(name, price_purchase, sku, category)')
      .eq('supplier_id', id)
      .order('created_at', { ascending: true });

    if (error) throw error;

    // flatten: products JOIN 결과를 평탄화
    const catalog = (data || []).map((row: Record<string, unknown>) => {
      const product = row.products as Record<string, unknown> | null;
      return {
        id: row.id,
        supplier_id: row.supplier_id,
        product_id: row.product_id,
        order_name: row.order_name || '',
        features: row.features || '',
        product_name: product?.name || '',
        price_purchase: product?.price_purchase || 0,
        sku: product?.sku || '',
        category: product?.category || '',
      };
    });

    return NextResponse.json({ catalog });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/suppliers/[id]/catalog — 매입품목 추가 */
export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { product_ids } = await req.json() as { product_ids: string[] };
    if (!product_ids || product_ids.length === 0) {
      return NextResponse.json({ error: '제품을 선택해주세요' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const rows = product_ids.map((pid) => ({ supplier_id: id, product_id: pid }));
    const { error } = await db
      .from('supplier_product_catalog')
      .upsert(rows, { onConflict: 'supplier_id,product_id', ignoreDuplicates: true });

    if (error) throw error;
    return NextResponse.json({ success: true, added: product_ids.length });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/suppliers/[id]/catalog — 주문명/특징 수정 */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    await params; // consume params
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { catalog_id, order_name, features } = await req.json() as {
      catalog_id: string;
      order_name?: string;
      features?: string;
    };

    if (!catalog_id) return NextResponse.json({ error: 'catalog_id 필수' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const updates: Record<string, unknown> = {};
    if (order_name !== undefined) updates.order_name = order_name;
    if (features !== undefined) updates.features = features;

    const { error } = await db
      .from('supplier_product_catalog')
      .update(updates)
      .eq('id', catalog_id);

    if (error) throw error;
    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/suppliers/[id]/catalog — 매입품목 삭제 */
export async function DELETE(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { catalog_id } = await req.json() as { catalog_id: string };
    if (!catalog_id) return NextResponse.json({ error: 'catalog_id 필수' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { error } = await db
      .from('supplier_product_catalog')
      .delete()
      .eq('id', catalog_id);

    if (error) throw error;
    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
