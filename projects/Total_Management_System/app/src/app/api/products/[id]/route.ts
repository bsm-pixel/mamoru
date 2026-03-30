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

/** DELETE /api/products/[id] — 제품 삭제 (연관 데이터 없을 때만) */
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

    // 연관 데이터 체크
    const [serialsRes, saleItemsRes, contractItemsRes] = await Promise.all([
      db.from('product_serials').select('id', { count: 'exact', head: true }).eq('product_id', id),
      db.from('offline_sale_items').select('id', { count: 'exact', head: true }).eq('product_id', id),
      db.from('contract_items').select('id', { count: 'exact', head: true }).eq('product_id', id),
    ]);

    const linked: string[] = [];
    if ((serialsRes.count || 0) > 0) linked.push(`시리얼 ${serialsRes.count}개`);
    if ((saleItemsRes.count || 0) > 0) linked.push(`판매 ${saleItemsRes.count}건`);
    if ((contractItemsRes.count || 0) > 0) linked.push(`계약서 ${contractItemsRes.count}건`);

    if (linked.length > 0) {
      return NextResponse.json({
        error: `연결된 데이터가 있어 삭제할 수 없습니다 (${linked.join(', ')}). 비활성화를 사용하세요.`,
      }, { status: 400 });
    }

    // 삭제
    const { error } = await db.from('products').delete().eq('id', id);
    if (error) throw error;

    return NextResponse.json({ deleted: true });
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

    const allowed = ['name', 'category', 'price', 'price_dealer', 'price_academy', 'price_purchase', 'supplier_id', 'description', 'imweb_product_no', 'barcode', 'image_url', 'is_active', 'product_group'];
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
