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

/** 이름 마스킹: 김미영 → 김*영 */
function maskName(name: string): string {
  if (!name) return '';
  if (name.length <= 1) return name;
  if (name.length === 2) return name[0] + '*';
  return name[0] + '*'.repeat(name.length - 2) + name[name.length - 1];
}

/** GET /api/reviews/info?uid=XXX&type=consult — 리뷰 폼 pre-fill 정보 (비인증) */
export async function GET(req: NextRequest) {
  try {
    const uid = req.nextUrl.searchParams.get('uid');
    const rawType = req.nextUrl.searchParams.get('type');
    // 'as'는 'repair'의 별칭으로 정규화 (솔라피 측 정적 URL이 'as'로 박혀있는 경우 방어망)
    const type = rawType === 'as' ? 'repair' : rawType;

    if (!uid || !type) {
      return NextResponse.json(
        { error: '잘못된 요청입니다' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // purchase 타입은 제품별 리뷰라 중복 체크를 items 단위로 수행 (아래 분기에서)
    if (type !== 'purchase') {
      const { data: existing } = await dbAny
        .from('reviews')
        .select('id')
        .eq('source_id', uid)
        .limit(1);

      if (existing && existing.length > 0) {
        return NextResponse.json(
          { error: '이미 후기를 작성하셨습니다', code: 'ALREADY_SUBMITTED' },
          { status: 409, headers: CORS_HEADERS }
        );
      }
    }

    if (type === 'purchase') {
      const { data: order } = await dbAny
        .from('orders')
        .select('id, orderer_name, ordered_at')
        .eq('id', uid)
        .single();

      if (!order) {
        return NextResponse.json(
          { error: '주문 정보를 찾을 수 없습니다' },
          { status: 404, headers: CORS_HEADERS }
        );
      }

      const { data: items } = await dbAny
        .from('order_items')
        .select('id, imweb_product_no, product_name, option_text, quantity')
        .eq('order_id', uid);

      // 이미 리뷰 작성된 제품 확인
      const { data: existingReviews } = await dbAny
        .from('reviews')
        .select('source_id')
        .like('source_id', `${uid}:%`);

      const reviewedProducts = new Set(
        (existingReviews || []).map((r: { source_id: string }) => r.source_id.split(':')[1])
      );

      // 전부 작성 완료 체크
      const allItems = items || [];
      const allReviewed = allItems.length > 0 && allItems.every(
        (it: { imweb_product_no: string }) => reviewedProducts.has(it.imweb_product_no)
      );

      if (allReviewed) {
        return NextResponse.json(
          { error: '모든 제품의 후기를 작성하셨습니다', code: 'ALREADY_SUBMITTED' },
          { status: 409, headers: CORS_HEADERS }
        );
      }

      return NextResponse.json({
        name: maskName(order.orderer_name),
        typeLabel: '제품구매',
        date: order.ordered_at ? new Date(order.ordered_at).toLocaleDateString('ko-KR') : '',
        items: allItems.map((it: { imweb_product_no: string; product_name: string; option_text: string | null; quantity: number }) => ({
          imweb_product_no: it.imweb_product_no,
          product_name: it.product_name,
          option_text: it.option_text,
          quantity: it.quantity,
          reviewed: reviewedProducts.has(it.imweb_product_no),
        })),
      }, { headers: CORS_HEADERS });
    }

    if (type === 'consult') {
      const { data: consult } = await dbAny
        .from('consultations')
        .select('name, consultation_type, visit_date, status')
        .eq('unique_id', uid)
        .single();

      if (!consult) {
        return NextResponse.json(
          { error: '상담 정보를 찾을 수 없습니다' },
          { status: 404, headers: CORS_HEADERS }
        );
      }

      const typeLabel = consult.consultation_type === 'store_visit' ? '매장 방문'
        : consult.consultation_type === 'field_request' ? '출장 상담' : '온라인상담';

      return NextResponse.json({
        name: maskName(consult.name),
        typeLabel,
        date: consult.visit_date || '',
        consultationType: consult.consultation_type,
      }, { headers: CORS_HEADERS });
    }

    if (type === 'repair') {
      const { data: repair } = await dbAny
        .from('repairs')
        .select('name, status')
        .eq('as_id', uid)
        .single();

      if (repair) {
        return NextResponse.json({
          name: maskName(repair.name),
          typeLabel: '복원수리',
        }, { headers: CORS_HEADERS });
      }
      // repairs에 없으면 → 아래 offline_sales fallback으로 진행
    }

    // 판매 건 fallback: uid가 판매번호(OS-*)인 경우 offline_sales에서 조회
    {
      const { data: sale } = await dbAny
        .from('offline_sales')
        .select('customer_name, sale_date, sale_number')
        .eq('sale_number', uid)
        .single();

      if (sale) {
        const typeLabels: Record<string, string> = { consult: '상담', repair: '복원수리', purchase: '제품구매' };
        return NextResponse.json({
          name: maskName(sale.customer_name),
          typeLabel: typeLabels[type] || type,
          date: sale.sale_date || '',
        }, { headers: CORS_HEADERS });
      }
    }

    return NextResponse.json(
      { error: '정보를 찾을 수 없습니다' },
      { status: 404, headers: CORS_HEADERS }
    );
  } catch (err) {
    console.error('[reviews/info] 조회 실패:', err);
    return NextResponse.json(
      { error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
