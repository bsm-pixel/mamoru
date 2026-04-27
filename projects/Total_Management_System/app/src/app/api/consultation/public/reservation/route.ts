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

/** GET /api/consultation/public/reservation?uid=xxx — 예약 정보 조회 (비인증)
 *  page_change_request.html에서 사용 (GAS getReservationInfo 대체) */
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
      .from('consultations')
      .select('id, unique_id, name, phone, consultation_type, status, visit_date, visit_time, address_road, address_detail, memo')
      .eq('unique_id', uid)
      .single();

    if (error || !data) {
      return NextResponse.json({ ok: false, error: '예약 정보를 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
    }

    const typeLabel: Record<string, string> = {
      store_visit: '매장 방문',
      field_request: '출장 요청',
      talk_consult: '톡 상담',
    };

    // 변경/취소 요청을 받을 수 있는 상태인지 (page_change_request.html이 분기 처리)
    const ALLOWED_STATUSES = ['confirmed', 'assigned'];
    const canRequestChange = ALLOWED_STATUSES.includes(String(data.status));

    return NextResponse.json({
      ok: true,
      canRequestChange,
      data: {
        uid: data.unique_id,
        name: data.name,
        phone: data.phone,
        type: typeLabel[data.consultation_type] || data.consultation_type,
        consultationType: data.consultation_type,
        status: data.status,
        date: data.visit_date || '',
        time: data.visit_time || '',
        address: [data.address_road, data.address_detail].filter(Boolean).join(' '),
        memo: data.memo || '',
      },
    }, { headers: CORS_HEADERS });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
