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

/** GET /api/reviews/public?type=all — 승인된 리뷰 공개 조회 (비인증)
 *  GET /api/reviews/public?group=R4 — 제품군별 리뷰 + 통계 (아임웹 상품 페이지용)
 *  GET /api/reviews/public?imweb_no=63 — 아임웹 상품번호로 자동 감지 (코드위젯 공통 삽입용) */
export async function GET(req: NextRequest) {
  try {
    const type = req.nextUrl.searchParams.get('type') || 'all';
    const group = req.nextUrl.searchParams.get('group'); // 제품군 필터
    const imwebNo = req.nextUrl.searchParams.get('imweb_no'); // 아임웹 상품번호 자동 감지

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // imweb_no → product_group 자동 resolve
    let resolvedGroup: string | null = group || null;
    if (!resolvedGroup && imwebNo) {
      const { data: prod } = await dbAny
        .from('products')
        .select('product_group, imweb_product_no')
        .eq('imweb_product_no', imwebNo)
        .single();
      if (prod?.product_group) {
        resolvedGroup = prod.product_group;
      } else if (prod) {
        // product_group 미설정 → 해당 상품 단독 리뷰 조회
        resolvedGroup = '__imweb_no__' + imwebNo;
      }
      // prod 없으면 → resolvedGroup null → 빈 결과 반환
    }

    let query = dbAny
      .from('reviews')
      .select('review_id, type, subtype, name, stars, content, photo_urls, created_at, product, product_group, meta, source, is_best')
      .eq('status', 'approved')
      .order('created_at', { ascending: false });

    // 필터 적용
    if (resolvedGroup && resolvedGroup.startsWith('__imweb_no__')) {
      // product_group 미설정 상품 → imweb_product_no로 직접 매칭
      const no = resolvedGroup.replace('__imweb_no__', '');
      query = query.eq('meta->>imweb_product_no', no);
    } else if (resolvedGroup) {
      query = query.eq('product_group', resolvedGroup);
    } else if (imwebNo) {
      // 상품을 못 찾은 경우 → 빈 결과
      return NextResponse.json({ reviews: [], stats: { count: 0, average: 0 } }, { headers: CORS_HEADERS });
    } else if (type !== 'all') {
      query = query.eq('type', type);
    }

    const { data, error } = await query.limit(200);
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

    // group 또는 imweb_no 파라미터 사용 시: { reviews, stats } 응답 (아임웹 제품 위젯용)
    if (group || imwebNo) {
      const reviews = data || [];
      const totalStars = reviews.reduce((sum: number, r: { stars: number }) => sum + r.stars, 0);
      const stats = {
        count: reviews.length,
        average: reviews.length > 0 ? Math.round((totalStars / reviews.length) * 10) / 10 : 0,
      };
      return NextResponse.json({ reviews, stats }, { headers: CORS_HEADERS });
    }

    // 기존 호환: flat array 응답
    return NextResponse.json(data || [], { headers: CORS_HEADERS });
  } catch (err) {
    console.error('[reviews/public] 조회 실패:', err);
    return NextResponse.json(
      { error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
