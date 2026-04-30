import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/reviews/promised — 약속 받았지만 아직 미작성 고객 통합 리스트
 *  3 source UNION: consultations + repairs + offline_sales
 *  조건: review_promised_at IS NOT NULL AND review_submitted_at IS NULL
 */
export async function GET() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 후기 약속/발송은 sale + repair source만 (consultation은 더 이상 리뷰 진입점 X — 2026-04-30 IA 정리)
    const [repairRes, saleRes] = await Promise.all([
      db.from('repairs')
        .select('id, as_id, name, phone, review_promised_at, review_request_sent_at')
        .not('review_promised_at', 'is', null)
        .is('review_submitted_at', null)
        .order('review_promised_at', { ascending: false })
        .limit(100),
      db.from('offline_sales')
        .select('id, sale_number, customer_name, customer_phone, review_promised_at, review_requested_at')
        .not('review_promised_at', 'is', null)
        .is('review_submitted_at', null)
        .is('cancelled_at', null)
        .order('review_promised_at', { ascending: false })
        .limit(100),
    ]);

    type Item = {
      source: 'repair' | 'sale';
      id: string;
      displayId: string;
      customerName: string;
      customerPhone: string | null;
      typeLabel: string;
      promisedAt: string;
      requestSentAt: string | null;
    };

    const items: Item[] = [];

    for (const r of (repairRes.data || [])) {
      items.push({
        source: 'repair',
        id: r.id,
        displayId: r.as_id,
        customerName: r.name,
        customerPhone: r.phone,
        typeLabel: '복원수리',
        promisedAt: r.review_promised_at,
        requestSentAt: r.review_request_sent_at,
      });
    }
    for (const s of (saleRes.data || [])) {
      items.push({
        source: 'sale',
        id: s.id,
        displayId: s.sale_number,
        customerName: s.customer_name,
        customerPhone: s.customer_phone,
        typeLabel: '판매',
        promisedAt: s.review_promised_at,
        requestSentAt: s.review_requested_at,
      });
    }

    // 약속 시각 최신순 정렬
    items.sort((a, b) => (a.promisedAt < b.promisedAt ? 1 : -1));

    return NextResponse.json({ items });
  } catch (err) {
    console.error('[reviews/promised] Error:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
