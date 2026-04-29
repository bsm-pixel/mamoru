import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/reviews/auto-match
 *  body: { source: 'consultation' | 'repair' | 'sale', id: string }
 *
 *  과거 작성된 리뷰가 067 배포 이전이라 review_submitted_at이 누락되었거나
 *  source_id 매칭 실패 등으로 누락된 경우, 카드 마운트 시 자동으로 reviews
 *  테이블을 검사하여 review_submitted_at을 백필.
 *
 *  매칭 기준 (source_id 정확 매칭):
 *  - consultation: reviews.source_id = consultations.unique_id AND type='consult'
 *  - repair:       reviews.source_id = repairs.as_id AND type='repair'
 *  - sale:         reviews.source_id = offline_sales.sale_number (type 무관)
 *
 *  멱등: review_submitted_at 이미 set이면 즉시 반환.
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const source = body.source as 'consultation' | 'repair' | 'sale';
    const id = body.id as string;

    if (!['consultation', 'repair', 'sale'].includes(source)) {
      return NextResponse.json({ error: 'invalid source' }, { status: 400 });
    }
    if (!id) return NextResponse.json({ error: 'id required' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // source 테이블 schema
    const tableMap: Record<typeof source, { table: string; sourceIdField: string; typeFilter?: 'consult' | 'repair' }> = {
      consultation: { table: 'consultations', sourceIdField: 'unique_id', typeFilter: 'consult' },
      repair: { table: 'repairs', sourceIdField: 'as_id', typeFilter: 'repair' },
      sale: { table: 'offline_sales', sourceIdField: 'sale_number' }, // type 무관
    };
    const cfg = tableMap[source];

    // 1) source row 조회 — review_submitted_at 이미 있으면 skip
    const { data: row, error: rowErr } = await db
      .from(cfg.table)
      .select(`id, review_submitted_at, ${cfg.sourceIdField}`)
      .eq('id', id)
      .single();
    if (rowErr || !row) {
      return NextResponse.json({ error: '대상을 찾을 수 없습니다' }, { status: 404 });
    }
    if (row.review_submitted_at) {
      return NextResponse.json({ matched: true, alreadySet: true, submittedAt: row.review_submitted_at });
    }

    const sourceIdValue = row[cfg.sourceIdField];
    if (!sourceIdValue) return NextResponse.json({ matched: false });

    // 2) reviews 테이블에서 매칭 검사
    let q = db.from('reviews')
      .select('id, review_id, type, source_id, created_at')
      .eq('source_id', sourceIdValue)
      .order('created_at', { ascending: true })
      .limit(1);
    if (cfg.typeFilter) q = q.eq('type', cfg.typeFilter);
    const { data: rev } = await q;

    if (!rev || rev.length === 0) {
      return NextResponse.json({ matched: false });
    }

    const matchedReview = rev[0];

    // 3) source 테이블에 review_submitted_at 백필 (reviews.created_at 사용)
    const submittedAt = matchedReview.created_at as string;
    await db
      .from(cfg.table)
      .update({ review_submitted_at: submittedAt })
      .eq('id', id)
      .is('review_submitted_at', null); // 멱등성 보강

    return NextResponse.json({
      matched: true,
      submittedAt,
      reviewId: matchedReview.review_id,
    });
  } catch (err) {
    console.error('[reviews/auto-match] Error:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
