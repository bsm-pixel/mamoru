import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * B2B 납품처(dealer/academy)별 납품품목 카탈로그 API
 *
 * supplier_product_catalog와 mirror 패턴 (마이그 073).
 * 다른 점: order_name(매입) → delivery_name(납품), products.supplier_id 자동 설정 X
 */

/** GET /api/customers/[id]/catalog — 납품품목 조회 */
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
      .from('customer_product_catalog')
      .select('id, customer_id, product_id, delivery_name, features, sort_order, created_at, products:product_id(name, price, sku, category, product_group, sort_order)')
      .eq('customer_id', id)
      .order('sort_order', { ascending: true })
      .order('created_at', { ascending: true });

    if (error) throw error;

    const catalog = (data || []).map((row: Record<string, unknown>) => {
      const product = row.products as Record<string, unknown> | null;
      return {
        id: row.id,
        customer_id: row.customer_id,
        product_id: row.product_id,
        delivery_name: row.delivery_name || '',
        features: row.features || '',
        sort_order: row.sort_order || 0,
        product_name: product?.name || '',
        price: product?.price || 0,
        sku: product?.sku || '',
        category: product?.category || '',
        product_group: product?.product_group || '',
      };
    });

    return NextResponse.json({ catalog });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/customers/[id]/catalog — 납품품목 추가 (다중) */
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

    const rows = product_ids.map((pid) => ({ customer_id: id, product_id: pid }));
    const { error } = await db
      .from('customer_product_catalog')
      .upsert(rows, { onConflict: 'customer_id,product_id', ignoreDuplicates: true });

    if (error) throw error;
    return NextResponse.json({ success: true, added: product_ids.length });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/customers/[id]/catalog — 납품명/특징 수정 (단건) */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { catalog_id, delivery_name, features } = await req.json() as {
      catalog_id: string;
      delivery_name?: string;
      features?: string;
    };

    if (!catalog_id) return NextResponse.json({ error: 'catalog_id 필수' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const updates: Record<string, unknown> = { updated_at: new Date().toISOString() };
    if (delivery_name !== undefined) updates.delivery_name = delivery_name;
    if (features !== undefined) updates.features = features;

    const { error } = await db
      .from('customer_product_catalog')
      .update(updates)
      .eq('id', catalog_id);

    if (error) throw error;
    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/customers/[id]/catalog — 납품품목 삭제 (단건) */
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
      .from('customer_product_catalog')
      .delete()
      .eq('id', catalog_id);

    if (error) throw error;
    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
