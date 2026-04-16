import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/products — 제품 목록 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const url = req.nextUrl;

    // SKU 중복 체크 (빠른 응답)
    const skuCheck = url.searchParams.get('sku_check');
    if (skuCheck) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data } = await (supabase as any).from('products').select('id').eq('sku', skuCheck).limit(1);
      return NextResponse.json({ exists: (data?.length || 0) > 0 });
    }

    const category = url.searchParams.get('category');
    const search = url.searchParams.get('search');
    const activeOnly = url.searchParams.get('active') !== 'false';

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    let query = db
      .from('products')
      .select('*')
      .order('product_group', { ascending: true, nullsFirst: false })
      .order('sort_order', { ascending: true })
      .order('name');

    if (activeOnly) query = query.eq('is_active', true);
    if (category) query = query.eq('category', category);
    if (search) {
      query = query.or(`name.ilike.%${search}%,sku.ilike.%${search}%`);
    }

    const { data, error } = await query;
    if (error) throw error;

    return NextResponse.json({ products: data || [] });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/products — 제품 등록 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { sku, name, category, price, price_dealer, price_academy, price_purchase, price_groups, supplier_id, description, imweb_product_no, barcode, image_url } = body;

    if (!sku?.trim() || !name?.trim()) {
      return NextResponse.json({ error: 'SKU와 제품명은 필수입니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: product, error } = await db
      .from('products')
      .insert({
        sku: sku.trim(),
        name: name.trim(),
        category: category || '',
        price: price || 0,
        price_dealer: price_dealer || 0, // dual-write (전환기)
        price_academy: price_academy || 0, // dual-write (전환기)
        price_purchase: price_purchase || 0,
        price_groups: price_groups || {},
        supplier_id: supplier_id || null,
        description: description || null,
        imweb_product_no: imweb_product_no || null,
        barcode: barcode || null,
        image_url: image_url || null,
      })
      .select()
      .single();

    if (error) throw error;

    return NextResponse.json({ product });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
