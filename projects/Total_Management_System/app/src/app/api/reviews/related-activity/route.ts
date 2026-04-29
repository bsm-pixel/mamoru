import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/reviews/related-activity
 *  query: ?phone=xxx&excludeSource=consultation|repair|sale&excludeId=xxx
 *
 *  같은 phone의 다른 source 리뷰 활동 조회 (자동 매칭 X — 정보 표시용).
 *  '활동 있음' 기준: review_promised_at OR review_request_sent_at OR review_submitted_at
 *  중 하나라도 NOT NULL.
 *
 *  자기 자신은 excludeSource + excludeId로 제외.
 *  최대 10건, 최신 활동(updated 기준) 우선.
 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const phone = req.nextUrl.searchParams.get('phone') || '';
    const excludeSource = req.nextUrl.searchParams.get('excludeSource') || '';
    const excludeId = req.nextUrl.searchParams.get('excludeId') || '';

    const phoneNorm = phone.replace(/\D/g, '');
    if (!phoneNorm) return NextResponse.json({ items: [] });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    type Row = { id: string; promisedAt: string | null; requestSentAt: string | null; submittedAt: string | null };
    type Item = {
      source: 'consultation' | 'repair' | 'sale';
      id: string;
      displayId: string;
      typeLabel: string;
      promisedAt: string | null;
      requestSentAt: string | null;
      submittedAt: string | null;
    };

    const items: Item[] = [];

    // 1) consultations — phone_normalized 사용
    if (excludeSource !== 'consultation' || true) {
      let q = db.from('consultations')
        .select('id, unique_id, consultation_type, review_promised_at, review_request_sent_at, review_submitted_at')
        .eq('phone_normalized', phoneNorm)
        .or('review_promised_at.not.is.null,review_request_sent_at.not.is.null,review_submitted_at.not.is.null')
        .limit(10);
      if (excludeSource === 'consultation' && excludeId) q = q.neq('id', excludeId);
      const { data: rows } = await q;
      for (const r of (rows || [])) {
        const typeLabel =
          r.consultation_type === 'store_visit' ? '매장상담'
          : r.consultation_type === 'field_request' ? '출장상담'
          : '톡상담';
        items.push({
          source: 'consultation',
          id: r.id,
          displayId: r.unique_id,
          typeLabel,
          promisedAt: r.review_promised_at,
          requestSentAt: r.review_request_sent_at,
          submittedAt: r.review_submitted_at,
        });
      }
    }

    // 2) repairs — phone_normalized 사용
    {
      let q = db.from('repairs')
        .select('id, as_id, review_promised_at, review_request_sent_at, review_submitted_at')
        .eq('phone_normalized', phoneNorm)
        .or('review_promised_at.not.is.null,review_request_sent_at.not.is.null,review_submitted_at.not.is.null')
        .limit(10);
      if (excludeSource === 'repair' && excludeId) q = q.neq('id', excludeId);
      const { data: rows } = await q;
      for (const r of (rows || [])) {
        items.push({
          source: 'repair',
          id: r.id,
          displayId: r.as_id,
          typeLabel: '복원수리',
          promisedAt: r.review_promised_at,
          requestSentAt: r.review_request_sent_at,
          submittedAt: r.review_submitted_at,
        });
      }
    }

    // 3) offline_sales — customer_phone 직접 정규화 매칭 (phone_normalized 컬럼 없음)
    {
      // customer_phone에 하이픈/공백 다양 → ilike로 부분 매칭은 부정확. 전수 조회 후 클라이언트 정규화 비교.
      let q = db.from('offline_sales')
        .select('id, sale_number, customer_phone, review_promised_at, review_requested_at, review_submitted_at, cancelled_at')
        .or('review_promised_at.not.is.null,review_requested_at.not.is.null,review_submitted_at.not.is.null')
        .is('cancelled_at', null)
        .limit(50); // 전체 활동 있는 판매 중 최대 50건만 가져와서 phone 매칭
      if (excludeSource === 'sale' && excludeId) q = q.neq('id', excludeId);
      const { data: rows } = await q;
      for (const r of (rows || [])) {
        const rowPhoneNorm = (r.customer_phone || '').replace(/\D/g, '');
        if (rowPhoneNorm !== phoneNorm) continue;
        items.push({
          source: 'sale',
          id: r.id,
          displayId: r.sale_number,
          typeLabel: '판매',
          promisedAt: r.review_promised_at,
          requestSentAt: r.review_requested_at, // semantic alias of review_request_sent_at
          submittedAt: r.review_submitted_at,
        });
      }
    }

    // 최신 활동 기준 정렬
    function latest(it: Item): string {
      return [it.submittedAt, it.requestSentAt, it.promisedAt].filter(Boolean).sort().reverse()[0] || '';
    }
    items.sort((a, b) => (latest(a) < latest(b) ? 1 : -1));

    return NextResponse.json({ items: items.slice(0, 10) });
  } catch (err) {
    console.error('[reviews/related-activity] Error:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
