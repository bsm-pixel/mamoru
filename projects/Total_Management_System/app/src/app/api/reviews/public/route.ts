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
      .select('review_id, type, subtype, name, stars, content, photo_urls, created_at, product, meta')
      .eq('status', 'approved')
      .order('created_at', { ascending: false });

    if (type !== 'all') {
      query = query.eq('type', type);
    }

    const { data, error } = await query.limit(50);
    if (error) throw error;

    return NextResponse.json(data || [], { headers: CORS_HEADERS });
  } catch (err) {
    console.error('[reviews/public] 조회 실패:', err);
    return NextResponse.json(
      { error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
