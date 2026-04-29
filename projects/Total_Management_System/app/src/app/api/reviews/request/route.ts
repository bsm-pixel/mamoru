import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { sendReviewRequestNotification, type ReviewSource } from '@/lib/notification/review-request';

/** POST /api/reviews/request — 후기 요청 알림톡 발송 (consultation / repair / sale 통합)
 *
 *  body 신규 형식: { source: 'consultation' | 'repair' | 'sale', id: string, review_type, subtype? }
 *  body 레거시:    { sale_id: string, review_type, subtype? }  ← 한 사이클 alias 유지 (회귀 방지)
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const review_type = body.review_type as 'consult' | 'repair' | 'purchase';
    const subtype = (body.subtype as string | undefined) || undefined;

    // 새 형식 / 레거시 호환
    let source: ReviewSource;
    let id: string;
    if (body.source && body.id) {
      source = body.source as ReviewSource;
      id = body.id as string;
    } else if (body.sale_id) {
      source = 'sale';
      id = body.sale_id as string;
    } else {
      return NextResponse.json({ error: 'source/id 또는 sale_id 필수' }, { status: 400 });
    }

    if (!['consultation', 'repair', 'sale'].includes(source)) {
      return NextResponse.json({ error: 'source는 consultation/repair/sale 중 하나' }, { status: 400 });
    }
    if (!review_type) {
      return NextResponse.json({ error: 'review_type 필수' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // source별 row 조회 + 정보 추출
    let sourceId: string;
    let customerName: string;
    let customerPhone: string | null;
    let updateTable: string;
    let updateField: string;

    if (source === 'consultation') {
      const { data: row, error } = await db
        .from('consultations')
        .select('id, unique_id, name, phone')
        .eq('id', id)
        .single();
      if (error || !row) return NextResponse.json({ error: '상담 건을 찾을 수 없습니다' }, { status: 404 });
      sourceId = row.unique_id;
      customerName = row.name;
      customerPhone = row.phone;
      updateTable = 'consultations';
      updateField = 'review_request_sent_at';
    } else if (source === 'repair') {
      const { data: row, error } = await db
        .from('repairs')
        .select('id, as_id, name, phone')
        .eq('id', id)
        .single();
      if (error || !row) return NextResponse.json({ error: '수리 건을 찾을 수 없습니다' }, { status: 404 });
      sourceId = row.as_id;
      customerName = row.name;
      customerPhone = row.phone;
      updateTable = 'repairs';
      updateField = 'review_request_sent_at';
    } else {
      // sale (offline_sales) — 레거시 review_requested_at 컬럼 사용
      const { data: row, error } = await db
        .from('offline_sales')
        .select('id, sale_number, customer_name, customer_phone')
        .eq('id', id)
        .single();
      if (error || !row) return NextResponse.json({ error: '판매 건을 찾을 수 없습니다' }, { status: 404 });
      sourceId = row.sale_number;
      customerName = row.customer_name;
      customerPhone = row.customer_phone;
      updateTable = 'offline_sales';
      updateField = 'review_requested_at';
    }

    if (!customerPhone) {
      return NextResponse.json({ error: '고객 연락처가 없어 알림톡을 발송할 수 없습니다' }, { status: 400 });
    }

    const result = await sendReviewRequestNotification({
      source,
      sourceId,
      customerName,
      customerPhone,
      reviewType: review_type,
      subtype,
    });

    // 발송 시각 기록
    await db.from(updateTable).update({ [updateField]: new Date().toISOString() }).eq('id', id);

    return NextResponse.json({ success: true, webhook_result: result });
  } catch (err) {
    console.error('[reviews/request] Error:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
