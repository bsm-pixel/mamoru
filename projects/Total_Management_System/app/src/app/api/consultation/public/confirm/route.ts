import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';
import { syncConsultationToCalendar } from '@/lib/google/calendar-sync';

const GITHUB_PAGES = 'bsm-pixel.github.io/mamoru/projects/consulting'; // Make 시나리오가 https:// 추가

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** GET /api/consultation/public/confirm?t=xxx&date=YYYY-MM-DD&time=HH:MM
 *  고객이 제안된 시간 중 하나를 선택하여 확정 (page_suggest.html에서 사용) */
export async function GET(req: NextRequest) {
  try {
    const token = req.nextUrl.searchParams.get('t');
    const date = req.nextUrl.searchParams.get('date');
    const time = req.nextUrl.searchParams.get('time');

    if (!token || !date || !time) {
      return NextResponse.json({ ok: false, error: 'token, date, time 필수' }, { status: 400, headers: CORS_HEADERS });
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // shortToken으로 조회
    const { data: list } = await dbAny
      .from('consultations')
      .select('id, unique_id, name, phone, consultation_type, status, gas_raw, address_road, address_detail')
      .eq('status', 'suggested');

    const data = (list || []).find((c: { gas_raw?: { shortToken?: string } }) =>
      c.gas_raw?.shortToken === token
    );

    if (!data) {
      return NextResponse.json({ ok: false, error: '제안 정보를 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
    }

    // 상태 변경 → confirmed + 날짜/시간 설정
    await dbAny.from('consultations').update({
      status: 'confirmed',
      visit_date: date,
      visit_time: time,
    }).eq('id', data.id);

    await dbAny.from('consultation_history').insert({
      consultation_id: data.id,
      from_status: 'suggested',
      to_status: 'confirmed',
      note: `고객 선택 확정: ${date} ${time}`,
    });

    // 확정 알림톡 — 다른 출장 알림톡(submit/suggest)과 동일하게 address·change_request_link 포함해야 함
    //   (이 두 변수가 빠지면 솔라피가 알림톡 발송 거부 → 문자 대체발송. 특히 change_request_link 는 "일정확인/변경" 버튼 URL)
    const phoneNorm = (data.phone || '').replace(/\D/g, '');
    const address = [data.address_road, data.address_detail].filter(Boolean).join(' ');
    try {
      await sendNotification({
        template: 'field_confirmed',
        phone: phoneNorm,
        name: data.name,
        data: {
          id: data.unique_id,
          name: data.name,
          phone: phoneNorm,
          type: '출장 요청',
          date,
          time,
          address,
          change_request_link: `${GITHUB_PAGES}/page_change_request.html?uid=${data.unique_id}`,
        },
      });
    } catch { /* 알림 실패해도 확정은 완료 */ }

    // 관리자 푸시 + Google Calendar 동기화 — after()로 응답 후 실행 보장 (Vercel 서버리스 대응)
    after(async () => {
      // 관리자 푸시 (고객 확정)
      try {
        const { sendPushToAll } = await import('@/lib/firebase/send-push');
        await sendPushToAll({
          title: '출장 일정 확정 ✅',
          body: `${data.name}님이 ${date} ${time}로 확정했습니다`,
          url: '/consultations',
          tag: `mamoru-confirm-${data.id}`,
          settingKey: 'push.field_confirmed',
        });
      } catch (e) {
        console.error('[confirm push] 실패:', e);
      }

      // Google Calendar 동기화
      try {
        await syncConsultationToCalendar(data.id);
      } catch (e) {
        console.error('[calendar-sync after confirm] 실패:', e);
      }
    });

    return NextResponse.json({ ok: true }, { headers: CORS_HEADERS });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
