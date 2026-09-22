import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

/** POST /api/consultation/delay — 출장 지연 안내 (직접 알림톡 발송) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const body = await req.json();
    const { consultationId, delayMin } = body as {
      consultationId: string;
      delayMin: number;
    };

    if (!consultationId || !delayMin || delayMin <= 0) {
      return NextResponse.json({ error: 'consultationId, delayMin(>0) 필수' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data: c, error: fetchErr } = await db
      .from('consultations')
      .select('id, unique_id, name, phone, consultation_type, status, visit_date, visit_time')
      .eq('id', consultationId)
      .single();

    if (fetchErr || !c) {
      return NextResponse.json({ error: '상담을 찾을 수 없습니다' }, { status: 404 });
    }

    if (c.status !== 'confirmed') {
      return NextResponse.json({ error: `확정 상태가 아닙니다 (현재: ${c.status})` }, { status: 400 });
    }
    if (c.consultation_type !== 'field_request') {
      return NextResponse.json({ error: '출장 예약만 지연 안내 가능합니다' }, { status: 400 });
    }

    // 도착 예정 시간 계산 — 방문 시간이 없으면 '원래 시간'이 비어 알림톡이 문자로 대체되므로 막는다
    const visitTime = String(c.visit_time || '').slice(0, 5);
    if (!/^\d{2}:\d{2}$/.test(visitTime)) {
      return NextResponse.json({ error: '방문 시간이 없어 지연 안내를 보낼 수 없습니다' }, { status: 400 });
    }
    const [h, m] = visitTime.split(':').map(Number);
    const revisedMin = h * 60 + m + delayMin;
    const revisedH = String(Math.floor(revisedMin / 60)).padStart(2, '0');
    const revisedM = String(revisedMin % 60).padStart(2, '0');
    const visitTimeRevised = `${revisedH}:${revisedM}`;

    // 알림톡 발송 — field_delayed
    const phoneNorm = (c.phone || '').replace(/\D/g, '');
    const sent = await sendNotification({
      template: 'field_delayed',
      phone: phoneNorm,
      name: c.name,
      data: {
        id: c.unique_id || c.id,
        name: c.name,
        phone: phoneNorm,
        type: '출장 요청',
        date: c.visit_date || '',
        time: visitTime,
        // 🚨 솔라피 템플릿 변수는 #{visit_time} — 'time' 만 보내면 변수 누락으로 문자 대체됐다 (2026-09-22 수정)
        visit_time: visitTime,
        delay_min: String(delayMin),
        visit_time_revised: visitTimeRevised,
      },
    });

    // 발송 실패(웹훅 미설정·Make 오류)는 성공으로 숨기지 않는다
    if (!sent.success) {
      return NextResponse.json({ error: `알림톡 발송 실패: ${sent.error || '알 수 없음'}` }, { status: 502 });
    }

    // 이력 기록
    await db.from('consultation_history').insert({
      consultation_id: consultationId,
      from_status: c.status,
      to_status: c.status,
      changed_by: user.id,
      note: `출장 지연 안내: ${delayMin}분 (도착 예정 ${visitTimeRevised})`,
    });

    return NextResponse.json({
      success: true,
      delay_min: delayMin,
      visit_time_revised: visitTimeRevised,
    });
  } catch (err) {
    console.error('[delay] 출장 지연 안내 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
