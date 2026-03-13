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
    const type = req.nextUrl.searchParams.get('type');

    if (!uid || !type) {
      return NextResponse.json(
        { error: '잘못된 요청입니다' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 이미 제출된 리뷰 확인
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
        : consult.consultation_type === 'field_request' ? '출장 상담' : '톡상담';

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
        .select('name, scissor_brand, status')
        .eq('as_id', uid)
        .single();

      if (!repair) {
        return NextResponse.json(
          { error: '복원수리 정보를 찾을 수 없습니다' },
          { status: 404, headers: CORS_HEADERS }
        );
      }

      return NextResponse.json({
        name: maskName(repair.name),
        typeLabel: '복원수리',
        scissorBrand: repair.scissor_brand || '',
      }, { headers: CORS_HEADERS });
    }

    return NextResponse.json(
      { error: '지원하지 않는 유형입니다' },
      { status: 400, headers: CORS_HEADERS }
    );
  } catch (err) {
    console.error('[reviews/info] 조회 실패:', err);
    return NextResponse.json(
      { error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
