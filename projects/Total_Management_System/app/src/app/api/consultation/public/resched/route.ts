import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendAdminEmail } from '@/lib/notification/email';
import { syncConsultationToCalendar } from '@/lib/google/calendar-sync';
import { errMsg } from '@/lib/utils/err';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** GET /api/consultation/public/resched?t=xxx&reason=yyy
 *  고객이 다른 일정 요청 (page_suggest.html에서 사용, GAS markResched 대체) */
export async function GET(req: NextRequest) {
  try {
    const token = req.nextUrl.searchParams.get('t');
    const reason = req.nextUrl.searchParams.get('reason') || '';

    if (!token) {
      return NextResponse.json({ ok: false, error: 'token 필수' }, { status: 400, headers: CORS_HEADERS });
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // shortToken으로 조회
    const { data: list } = await dbAny
      .from('consultations')
      .select('id, unique_id, name, phone, consultation_type, status, memo, gas_raw')
      .in('status', ['suggested', 'confirmed']);

    const data = (list || []).find((c: { gas_raw?: { shortToken?: string } }) =>
      c.gas_raw?.shortToken === token
    );

    if (!data) {
      return NextResponse.json({ ok: false, error: '예약 정보를 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
    }

    // 상태 변경 → reschedule_requested + 메모에 재요청 사유 추가
    const newMemo = data.memo
      ? `${data.memo}\n[고객 재요청 ${new Date().toLocaleString('ko-KR')}] ${reason}`
      : `[고객 재요청 ${new Date().toLocaleString('ko-KR')}] ${reason}`;

    // gas_raw에 재요청 사유 저장 (캘린더 이벤트 설명에 표시)
    const mergedGasRaw = {
      ...(data.gas_raw || {}),
      reschedule_reason: reason || '',
      reschedule_requested_at: new Date().toISOString(),
    };

    await dbAny.from('consultations').update({
      status: 'reschedule_requested',
      memo: newMemo,
      gas_raw: mergedGasRaw,
    }).eq('id', data.id);

    await dbAny.from('consultation_history').insert({
      consultation_id: data.id,
      from_status: data.status,
      to_status: 'reschedule_requested',
      note: reason ? `고객 재요청: ${reason}` : '고객 일정 재요청',
    });

    // 관리자 이메일 알림
    try {
      await sendAdminEmail(
        `[MAMORU] 출장 일정 재요청 — ${data.name}`,
        `고객: ${data.name}\n연락처: ${data.phone}\n사유: ${reason || '(없음)'}\n\nTMS에서 새 시간을 제안해주세요.`
      );
    } catch { /* 이메일 실패해도 재요청은 완료 */ }

    // Google Calendar 동기화 — reschedule_requested 상태로 이벤트 ⏳ 접두어 업데이트 (즉시)
    try {
      await syncConsultationToCalendar(data.id);
    } catch (e) {
      console.error('[calendar-sync resched] 실패:', { id: data.id, error: errMsg(e) });
    }

    // 관리자 푸시는 응답 후 실행 (지연 허용)
    after(async () => {
      try {
        const { sendPushToAll } = await import('@/lib/firebase/send-push');
        await sendPushToAll({
          title: '출장 일정 재요청 🔄',
          body: reason
            ? `${data.name}님 재요청 — ${reason.slice(0, 50)}${reason.length > 50 ? '...' : ''}`
            : `${data.name}님이 다른 시간을 요청했습니다`,
          url: '/consultations',
          tag: `mamoru-resched-${data.id}`,
          settingKey: 'push.field_reschedule',
        });
      } catch (e) {
        console.error('[resched push] 실패:', e);
      }
    });

    return NextResponse.json({ ok: true }, { headers: CORS_HEADERS });
  } catch (err) {
    return NextResponse.json({ ok: false, error: errMsg(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
