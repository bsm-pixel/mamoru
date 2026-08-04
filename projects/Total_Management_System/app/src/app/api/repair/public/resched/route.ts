import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};
const GITHUB_PAGES = 'page.mamoru.kr/projects/as'; // Make 시나리오가 https:// 추가

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** 'YYYY-MM-DD' → '7월 30일 (수)' */
function formatKoreanDate(dateStr: string): string {
  const d = new Date(`${dateStr}T00:00:00+09:00`);
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  return `${d.getMonth() + 1}월 ${d.getDate()}일 (${days[d.getDay()]})`;
}

/** POST /api/repair/public/resched — 직접방문 고객 셀프 일정변경 (즉시 반영 + as_visit_rescheduled)
 *  body: { uid(manage_token), visit_date:'YYYY-MM-DD', visit_time:'HH:MM' } */
export async function POST(req: NextRequest) {
  try {
    const { uid, visit_date, visit_time } = await req.json();
    if (!uid || !visit_date || !visit_time) {
      return NextResponse.json({ ok: false, error: 'uid·방문일·시간 필수' }, { status: 400, headers: CORS_HEADERS });
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const { data: r, error } = await dbAny
      .from('repairs')
      .select('id, as_id, manage_token, name, phone, proceed_type, status, qty_mamoru, qty_other')
      .eq('manage_token', uid)
      .single();
    if (error || !r) {
      return NextResponse.json({ ok: false, error: '예약을 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
    }
    if (r.proceed_type !== '직접방문' || r.status !== 'intake') {
      return NextResponse.json({ ok: false, error: '변경할 수 없는 예약 상태입니다' }, { status: 400, headers: CORS_HEADERS });
    }

    // 기본 중복 가드 — 같은 날짜·시간의 다른 활성 직접방문 건이 있으면 거절 (상세 충돌은 slots API가 페이지에서 처리)
    const { data: clash } = await dbAny
      .from('repairs')
      .select('id')
      .eq('proceed_type', '직접방문')
      .eq('visit_date', visit_date)
      .eq('visit_time', visit_time)
      .not('status', 'eq', 'cancelled')
      .neq('id', r.id)
      .limit(1);
    if (clash && clash.length > 0) {
      return NextResponse.json({ ok: false, error: '이미 예약된 시간입니다. 다른 시간을 선택해주세요.' }, { status: 409, headers: CORS_HEADERS });
    }

    const qty = (r.qty_mamoru || 0) + (r.qty_other || 0);
    const durationMin = 10 + (Math.max(qty, 1) - 1) * 5; // submit·slots 와 동일 공식(수량별 소요시간)

    const { error: updErr } = await dbAny
      .from('repairs')
      .update({
        visit_date,
        visit_time,
        visit_duration_min: durationMin,
        // 일정 변경 → 리마인드 재발송되도록 발송 플래그 리셋(옛 일정 리마인드 방지)
        visit_remind_24h_sent_at: null,
        visit_remind_2h_sent_at: null,
        admin_note: `[고객 셀프 일정변경] ${visit_date} ${visit_time}`,
        updated_at: new Date().toISOString(),
      })
      .eq('id', r.id);
    if (updErr) {
      return NextResponse.json({ ok: false, error: '변경 처리 실패' }, { status: 500, headers: CORS_HEADERS });
    }

    // 변경 안내 알림톡 (검수 통과 후 실제 발송 — 미등록이면 try/catch로 무시)
    try {
      await sendNotification({
        template: 'as_visit_rescheduled',
        phone: r.phone,
        name: r.name,
        data: {
          as_id: r.as_id,
          visit_date: formatKoreanDate(visit_date),
          visit_time,
          qty: String(qty),
          change_request_link: `${GITHUB_PAGES}/page_change_request.html?uid=${r.manage_token}`,
        },
      });
    } catch (e) {
      console.error('[repair/resched] 알림톡 발송 실패(변경은 완료):', e);
    }

    return NextResponse.json({ ok: true, visit_date, visit_time }, { headers: CORS_HEADERS });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
