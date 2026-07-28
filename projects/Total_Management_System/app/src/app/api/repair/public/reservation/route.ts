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

/** GET /api/repair/public/reservation?uid=<manage_token> — 직접방문 예약 정보 조회 (비인증, 토큰 기반)
 *  projects/as/page_change_request.html 에서 사용 (121, 2026-07-28) */
export async function GET(req: NextRequest) {
  try {
    const uid = req.nextUrl.searchParams.get('uid');
    if (!uid) {
      return NextResponse.json({ ok: false, error: 'uid 필수' }, { status: 400, headers: CORS_HEADERS });
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const { data, error } = await dbAny
      .from('repairs')
      .select('id, as_id, manage_token, name, phone, proceed_type, status, visit_date, visit_time, qty_mamoru, qty_other')
      .eq('manage_token', uid)
      .single();

    if (error || !data) {
      return NextResponse.json({ ok: false, error: '예약 정보를 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
    }

    // 직접방문 + 방문 전(intake) 상태에서만 셀프 변경/취소 허용 (수리 시작/완료/취소 후엔 불가)
    const canRequestChange = data.proceed_type === '직접방문' && data.status === 'intake';
    const qty = (data.qty_mamoru || 0) + (data.qty_other || 0);

    return NextResponse.json({
      ok: true,
      canRequestChange,
      data: {
        uid: data.manage_token,
        asId: data.as_id,
        name: data.name,
        proceedType: data.proceed_type,
        status: data.status,
        date: data.visit_date || '',
        time: data.visit_time || '',
        qty,
      },
    }, { headers: CORS_HEADERS });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
