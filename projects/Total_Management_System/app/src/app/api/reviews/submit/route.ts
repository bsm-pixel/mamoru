import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** 리뷰 ID 자동 채번: RV-YYYYMMDD-NNN */
async function generateReviewId(db: ReturnType<typeof createServiceClient>): Promise<string> {
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const prefix = `RV-${today}-`;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data } = await (db as any)
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

/** POST /api/reviews/submit — 고객 리뷰 제출 (비인증, 공개 엔드포인트) */
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { uid, type, stars, content, photoUrls } = body;

    if (!uid || !type || !stars || !content) {
      return NextResponse.json(
        { error: '필수 항목을 입력해주세요' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    if (stars < 1 || stars > 5) {
      return NextResponse.json(
        { error: '별점은 1~5 사이여야 합니다' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    // Service Role 클라이언트 (RLS 우회하여 조회 가능)
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 중복 제출 방지: 같은 source_id로 이미 작성된 리뷰 확인
    const { data: existing } = await dbAny
      .from('reviews')
      .select('id')
      .eq('source_id', uid)
      .limit(1);

    if (existing && existing.length > 0) {
      return NextResponse.json(
        { error: '이미 후기를 작성하셨습니다' },
        { status: 409, headers: CORS_HEADERS }
      );
    }

    // uid로 원본 데이터 조회 (상담 또는 수리)
    let name = '';
    let phone = '';
    let subtype = '';
    let meta: Record<string, string> = {};

    if (type === 'consult') {
      const { data: consult } = await dbAny
        .from('consultations')
        .select('name, phone, consultation_type, visit_date, visit_time')
        .eq('unique_id', uid)
        .single();

      if (!consult) {
        return NextResponse.json(
          { error: '상담 정보를 찾을 수 없습니다' },
          { status: 404, headers: CORS_HEADERS }
        );
      }
      name = consult.name;
      phone = consult.phone || '';
      subtype = consult.consultation_type || '';
      meta = {
        visit_date: consult.visit_date || '',
        visit_time: consult.visit_time || '',
        consultation_type: consult.consultation_type || '',
      };
    } else if (type === 'repair') {
      const { data: repair } = await dbAny
        .from('repairs')
        .select('name, phone, proceed_type')
        .eq('as_id', uid)
        .single();

      if (!repair) {
        return NextResponse.json(
          { error: '복원수리 정보를 찾을 수 없습니다' },
          { status: 404, headers: CORS_HEADERS }
        );
      }
      name = repair.name;
      phone = repair.phone || '';
      subtype = 'restoration';
      meta = {
        proceed_type: repair.proceed_type || '',
      };
    } else {
      return NextResponse.json(
        { error: '지원하지 않는 리뷰 유형입니다' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const reviewId = await generateReviewId(db);

    const { data, error } = await dbAny
      .from('reviews')
      .insert({
        review_id: reviewId,
        type,
        subtype,
        name,
        phone,
        stars: Number(stars),
        content: String(content).trim(),
        photo_urls: Array.isArray(photoUrls) ? photoUrls : [],
        source_id: uid,
        status: 'pending',
        meta,
      })
      .select()
      .single();

    if (error) throw error;

    return NextResponse.json(
      { success: true, reviewId: data.review_id },
      { headers: CORS_HEADERS }
    );
  } catch (err) {
    console.error('[reviews/submit] 제출 실패:', err);
    return NextResponse.json(
      { error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
