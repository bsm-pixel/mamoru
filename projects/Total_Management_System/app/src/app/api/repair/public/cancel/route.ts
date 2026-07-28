import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

function formatKoreanDate(dateStr: string): string {
  if (!dateStr) return '';
  const d = new Date(`${dateStr}T00:00:00+09:00`);
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  return `${d.getMonth() + 1}월 ${d.getDate()}일 (${days[d.getDay()]})`;
}

/** POST /api/repair/public/cancel — 직접방문 고객 셀프 취소 (status=cancelled + as_visit_cancelled)
 *  body: { uid(manage_token), reason? } */
export async function POST(req: NextRequest) {
  try {
    const { uid, reason } = await req.json();
    if (!uid) {
      return NextResponse.json({ ok: false, error: 'uid 필수' }, { status: 400, headers: CORS_HEADERS });
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const { data: r, error } = await dbAny
      .from('repairs')
      .select('id, as_id, name, phone, proceed_type, status, visit_date, visit_time')
      .eq('manage_token', uid)
      .single();
    if (error || !r) {
      return NextResponse.json({ ok: false, error: '예약을 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
    }
    if (r.proceed_type !== '직접방문' || r.status !== 'intake') {
      return NextResponse.json({ ok: false, error: '취소할 수 없는 예약 상태입니다' }, { status: 400, headers: CORS_HEADERS });
    }

    const { error: updErr } = await dbAny
      .from('repairs')
      .update({
        status: 'cancelled',
        admin_note: `[고객 셀프 취소]${reason ? ' ' + String(reason).slice(0, 200) : ''}`,
        updated_at: new Date().toISOString(),
      })
      .eq('id', r.id);
    if (updErr) {
      return NextResponse.json({ ok: false, error: '취소 처리 실패' }, { status: 500, headers: CORS_HEADERS });
    }

    // 취소 완료 알림톡 (검수 통과 후 실제 발송)
    try {
      await sendNotification({
        template: 'as_visit_cancelled',
        phone: r.phone,
        name: r.name,
        data: {
          as_id: r.as_id,
          visit_date: formatKoreanDate(r.visit_date),
          visit_time: r.visit_time || '',
        },
      });
    } catch (e) {
      console.error('[repair/cancel] 알림톡 발송 실패(취소는 완료):', e);
    }

    return NextResponse.json({ ok: true }, { headers: CORS_HEADERS });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
