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

/** GET /api/consultation/public/suggest-data?t=xxx — 시간 제안 데이터 조회 (비인증)
 *  page_suggest.html에서 사용 (GAS getSuggestData 대체) */
export async function GET(req: NextRequest) {
  try {
    const token = req.nextUrl.searchParams.get('t');
    if (!token) {
      return NextResponse.json({ ok: false, error: 'token 필수' }, { status: 400, headers: CORS_HEADERS });
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // shortToken으로 조회 (gas_raw.shortToken)
    const { data: list } = await dbAny
      .from('consultations')
      .select('id, unique_id, name, phone, consultation_type, status, visit_date, visit_time, address_road, address_detail, suggestions, gas_raw')
      .eq('status', 'suggested');

    // shortToken 매칭
    const data = (list || []).find((c: { gas_raw?: { shortToken?: string } }) =>
      c.gas_raw?.shortToken === token
    );

    if (!data) {
      return NextResponse.json({ ok: false, error: '제안 정보를 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
    }

    // suggestions 파싱
    const suggestions = data.suggestions?.dates || [];

    // GAS와 동일한 buttons 형식 (page_suggest.html parseSlots 호환)
    const API_CONFIRM = 'https://app-eta-sandy-75.vercel.app/api/consultation/public/confirm';
    const buttons = suggestions.map((s: { date: string; time: string }) => ({
      label: `${s.date} ${s.time}`,
      url: `${API_CONFIRM}?t=${encodeURIComponent(token)}&date=${encodeURIComponent(s.date)}&time=${encodeURIComponent(s.time)}`,
    }));

    return NextResponse.json({
      ok: true,
      data: {
        uid: data.unique_id,
        name: data.name,
        phone: data.phone,
        status: data.status,
        address: [data.address_road, data.address_detail].filter(Boolean).join(' '),
        buttons,
      },
    }, { headers: CORS_HEADERS });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
