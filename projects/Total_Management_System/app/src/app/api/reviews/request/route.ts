import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { sendNotification, type NotifyTemplate } from '@/lib/notification/make-webhook';

/** POST /api/reviews/request — 판매 건에서 후기 요청 알림톡 발송 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { sale_id, review_type, subtype } = body as {
      sale_id: string;
      review_type: 'consult' | 'repair' | 'purchase';
      subtype?: string; // store_visit, field_request, talk_consult
    };

    if (!sale_id || !review_type) {
      return NextResponse.json({ error: 'sale_id와 review_type은 필수입니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 판매 건 조회
    const { data: sale, error: saleError } = await db
      .from('offline_sales')
      .select('id, sale_number, customer_name, customer_phone, review_requested_at')
      .eq('id', sale_id)
      .single();

    if (saleError || !sale) {
      return NextResponse.json({ error: '판매 건을 찾을 수 없습니다' }, { status: 404 });
    }

    if (!sale.customer_phone) {
      return NextResponse.json({ error: '고객 연락처가 없어 알림톡을 발송할 수 없습니다' }, { status: 400 });
    }

    // 리뷰폼 URL 생성
    const reviewFormBase = 'https://bsm-pixel.github.io/mamoru/projects/reviews/page_review.html';
    const reviewUrl = `${reviewFormBase}?type=${review_type}&uid=${sale.sale_number}&name=${encodeURIComponent(sale.customer_name)}${subtype ? `&subtype=${subtype}` : ''}`;

    // 템플릿 선택
    let template: NotifyTemplate;
    if (review_type === 'repair') {
      template = 'as_review_request';
    } else if (review_type === 'purchase') {
      template = 'purchase_review_request';
    } else {
      template = 'review_request';
    }

    // Make webhook 발송
    const result = await sendNotification({
      template,
      phone: sale.customer_phone,
      name: sale.customer_name,
      data: {
        id: sale.sale_number,
        uid: sale.sale_number,
        review_type,
        type_label: review_type === 'repair' ? '복원수리' : review_type === 'consult' ? '상담' : '제품구매',
        subtype: subtype || '',
        review_url: reviewUrl,
      },
    });

    // review_requested_at 기록
    await db
      .from('offline_sales')
      .update({ review_requested_at: new Date().toISOString() })
      .eq('id', sale_id);

    return NextResponse.json({ success: true, webhook_result: result });
  } catch (err) {
    console.error('[reviews/request] Error:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
