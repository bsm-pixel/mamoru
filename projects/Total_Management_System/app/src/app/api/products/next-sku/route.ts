import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { computeNextSku } from '@/lib/product/next-sku';

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

    // 정확히 `{category}\d+` 형식만 인식하는 공통 채번 헬퍼 (유사접두어 오염·충돌 방지)
    const nextSku = await computeNextSku(db, category);
    const number = parseInt(nextSku.slice(category.length), 10) || 1;

    return NextResponse.json({ sku: nextSku, category, number });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
