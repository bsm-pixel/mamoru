import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

/** 리뷰 ID 자동 채번: RV-YYYYMMDD-NNN */
async function generateReviewId(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  dbAny: any,
  date?: string
): Promise<string> {
  const d = date ? new Date(date) : new Date();
  const today = d.toISOString().slice(0, 10).replace(/-/g, '');
  const prefix = `RV-${today}-`;

  const { data } = await dbAny
    .from('reviews')
    .select('review_id')
    .like('review_id', `${prefix}%`)
    .order('review_id', { ascending: false })
    .limit(1);

  let seq = 1;
  if (data && data.length > 0) {
    const last = data[0].review_id as string;
    seq = parseInt(last.split('-').pop() || '0', 10) + 1;
  }

  return `${prefix}${String(seq).padStart(3, '0')}`;
}

/** POST /api/reviews/naver — 네이버 리뷰 단건/일괄 등록 (인증 필요) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const body = await req.json();
    const { reviews } = body as { reviews: NaverReviewInput[] };

    if (!Array.isArray(reviews) || reviews.length === 0) {
      return NextResponse.json({ error: '등록할 리뷰가 없습니다' }, { status: 400 });
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const results: { index: number; review_id: string; error?: string }[] = [];

    for (let i = 0; i < reviews.length; i++) {
      const r = reviews[i];
      try {
        if (!r.name || !r.stars || !r.content || !r.type) {
          results.push({ index: i, review_id: '', error: '필수 항목 누락 (name, stars, content, type)' });
          continue;
        }

        const reviewId = await generateReviewId(dbAny, r.created_at);

        const { error } = await dbAny.from('reviews').insert({
          review_id: reviewId,
          type: r.type,
          subtype: r.subtype || null,
          name: r.name,
          phone: '',
          stars: Math.min(5, Math.max(1, Number(r.stars))),
          content: String(r.content).trim(),
          photo_urls: Array.isArray(r.photo_urls) ? r.photo_urls : [],
          source_id: `naver-${Date.now()}-${i}`,
          product: r.product || null,
          status: 'approved', // 네이버 리뷰는 이미 검증됨 → 바로 승인
          approved_at: new Date().toISOString(),
          source: 'naver',
          is_best: r.is_best || false,
          meta: {
            ...(r.meta || {}),
            naver_date: r.created_at || '',
            imported_at: new Date().toISOString(),
          },
          created_at: r.created_at || new Date().toISOString(),
        });

        if (error) throw error;
        results.push({ index: i, review_id: reviewId });
      } catch (err) {
        results.push({ index: i, review_id: '', error: String(err) });
      }
    }

    const success = results.filter(r => !r.error).length;
    const failed = results.filter(r => r.error).length;

    return NextResponse.json({ success, failed, results });
  } catch (err) {
    console.error('[reviews/naver] 등록 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

interface NaverReviewInput {
  type: 'consult' | 'repair' | 'purchase';
  subtype?: string;
  name: string;
  stars: number;
  content: string;
  photo_urls?: string[];
  product?: string;
  is_best?: boolean;
  created_at?: string;
  meta?: Record<string, string>;
}
