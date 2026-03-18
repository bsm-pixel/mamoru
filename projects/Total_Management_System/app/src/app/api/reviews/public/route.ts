import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** GET /api/reviews/public?type=all — 승인된 리뷰 공개 조회 (비인증) */
export async function GET(req: NextRequest) {
  try {
    const type = req.nextUrl.searchParams.get('type') || 'all';

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    let query = dbAny
      .from('reviews')
      .select('review_id, type, subtype, name, stars, content, photo_urls, created_at, product, meta, source, is_best')
      .eq('status', 'approved')
      .order('created_at', { ascending: false });

    if (type !== 'all') {
      query = query.eq('type', type);
    }

    const { data, error } = await query.limit(50);
    if (error) throw error;

    // purchase 리뷰에 제품 이미지 enrichment
    if (data && data.length > 0) {
      const purchaseReviews = data.filter(
        (r: { type: string; meta?: { imweb_product_no?: string } }) =>
          r.type === 'purchase' && r.meta?.imweb_product_no
      );
      if (purchaseReviews.length > 0) {
        const productNos = purchaseReviews.map(
          (r: { meta: { imweb_product_no: string } }) => r.meta.imweb_product_no
        );
        const { data: products } = await dbAny
          .from('products')
          .select('imweb_product_no, image_url')
          .in('imweb_product_no', productNos);

        const imageMap: Record<string, string> = {};
        (products || []).forEach((p: { imweb_product_no: string; image_url: string }) => {
          imageMap[p.imweb_product_no] = p.image_url;
        });

        data.forEach((r: { type: string; meta?: { imweb_product_no?: string }; product_image_url?: string | null }) => {
          if (r.type === 'purchase' && r.meta?.imweb_product_no) {
            r.product_image_url = imageMap[r.meta.imweb_product_no] || null;
          }
        });
      }
    }

    return NextResponse.json(data || [], { headers: CORS_HEADERS });
  } catch (err) {
    console.error('[reviews/public] 조회 실패:', err);
    return NextResponse.json(
      { error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
