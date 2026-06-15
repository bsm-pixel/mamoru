import { NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { stockLabel } from '@/lib/event/options';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** GET /api/event/public/products — EVENT 카탈로그 (비인증, CORS)
 *  category='EVENT' && is_active 품목 + 가격 + 재고라벨(적을 때만 수량) + tags(event_type/spec/dry_subtype) */
export async function GET() {
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (db as any)
      .from('products')
      .select('id, sku, name, category, price, stock_quantity, tags, is_active')
      .eq('category', 'EVENT')
      .eq('is_active', true)
      .order('price', { ascending: true });

    if (error) throw error;

    const items = (data || []).map((p: {
      id: string; sku: string; name: string; price: number; stock_quantity: number;
      tags: Record<string, unknown> | null;
    }) => {
      const stock = p.stock_quantity ?? 0;
      const label = stockLabel(stock);
      const tags = p.tags || {};
      return {
        id: p.id,
        sku: p.sku,
        name: p.name,
        price: p.price,
        soldout: stock <= 0,
        stock_text: label.text,
        stock_tone: label.tone,
        event_type: (tags.event_type as string) || '',  // blunt|thinning|long|dry
        spec: (tags.spec as string) || '',                // 인치/감모/DRY서브
        dry_subtype: (tags.dry_subtype as string) || '',
      };
    });

    return NextResponse.json({ ok: true, items }, { headers: CORS_HEADERS });
  } catch (err) {
    console.error('[event/public/products] 실패:', err);
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
