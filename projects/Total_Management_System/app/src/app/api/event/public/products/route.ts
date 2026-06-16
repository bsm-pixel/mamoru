import { NextRequest, NextResponse } from 'next/server';
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

/** GET /api/event/public/products?campaign=<id> — EVENT 카탈로그 (비인증, CORS)
 *  해당 캠페인 품목만(tags.campaign_id) + 가격 + 재고라벨 + 손/종류/옵션 tags */
export async function GET(req: NextRequest) {
  try {
    const campaign = req.nextUrl.searchParams.get('campaign');
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let q = (db as any)
      .from('products')
      .select('id, sku, name, category, price, stock_quantity, tags, is_active')
      .eq('category', 'EVENT')
      .eq('is_active', true);
    // 캠페인 스코프 — 그 캠페인 품목만 (앰버서더/할인 등 다른 캠페인 품목 섞임 방지)
    if (campaign) q = q.eq('tags->>campaign_id', campaign);
    const { data, error } = await q.order('price', { ascending: true });

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
        hand: (tags.hand as string) || 'right',           // right|left
        event_type: (tags.event_type as string) || '',    // blunt|thinning|long|dry
        spec: (tags.spec as string) || '',                 // 인치/감모/DRY서브
        dry_subtype: (tags.dry_subtype as string) || '',
      };
    });

    return NextResponse.json({ ok: true, items }, { headers: CORS_HEADERS });
  } catch (err) {
    console.error('[event/public/products] 실패:', err);
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
