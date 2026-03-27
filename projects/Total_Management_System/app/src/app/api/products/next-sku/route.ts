import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * GET /api/products/next-sku?category=BL
 * 해당 카테고리의 다음 SKU 자동 채번
 * 예: BL 카테고리에 BL001~BL005 있으면 → BL006 반환
 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const category = req.nextUrl.searchParams.get('category') || 'BL';

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 해당 카테고리의 SKU 중 가장 큰 번호 조회
    const { data: products } = await db
      .from('products')
      .select('sku')
      .like('sku', `${category}%`)
      .order('sku', { ascending: false })
      .limit(1);

    let nextNum = 1;
    if (products && products.length > 0) {
      const lastSku = products[0].sku; // 예: BL005
      const numPart = lastSku.replace(category, '');
      const parsed = parseInt(numPart);
      if (!isNaN(parsed)) nextNum = parsed + 1;
    }

    const nextSku = `${category}${String(nextNum).padStart(3, '0')}`;

    return NextResponse.json({ sku: nextSku, category, number: nextNum });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
