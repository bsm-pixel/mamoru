import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** GET /api/consultation/public/cancel?uid=xxx&reason=yyy — 고객 취소 요청 (비인증)
 *  page_change_request.html에서 사용 (GAS submitChangeRequest action=cancel 대체) */
export async function GET(req: NextRequest) {
  try {
    const uid = req.nextUrl.searchParams.get('uid');
    const reason = req.nextUrl.searchParams.get('reason') || '';
    if (!uid) {
      return NextResponse.json({ ok: false, error: 'uid 필수' }, { status: 400, headers: CORS_HEADERS });
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const { data, error } = await dbAny
      .from('consultations')
      .select('id, unique_id, name, phone, consultation_type, status')
      .eq('unique_id', uid)
      .single();

    if (error || !data) {
      return NextResponse.json({ ok: false, error: '예약을 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
    }

    if (['cancelled', 'completed'].includes(data.status)) {
      return NextResponse.json({ ok: false, error: '이미 취소/완료된 예약입니다' }, { status: 400, headers: CORS_HEADERS });
    }

    // 상태 변경
    await dbAny.from('consultations').update({ status: 'cancelled' }).eq('id', data.id);

    // 이력
    await dbAny.from('consultation_history').insert({
      consultation_id: data.id,
      from_status: data.status,
      to_status: 'cancelled',
      note: reason ? `고객 취소: ${reason}` : '고객 취소',
    });

    // 취소 알림톡
    const phoneNorm = (data.phone || '').replace(/\D/g, '');
    const template = data.consultation_type === 'field_request' ? 'field_cancelled' : 'cancelled';
    try {
      await sendNotification({
        template,
        phone: phoneNorm,
        name: data.name,
        data: {
          id: data.unique_id,
          name: data.name,
          phone: phoneNorm,
          type: data.consultation_type === 'store_visit' ? '매장 방문' : '출장 요청',
        },
      });
    } catch { /* 알림 실패해도 취소는 완료 */ }

    return NextResponse.json({ ok: true }, { headers: CORS_HEADERS });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
