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

    // 도착 예정 시간 계산
    const [h, m] = (c.visit_time || '00:00').split(':').map(Number);
    const revisedMin = h * 60 + m + delayMin;
    const revisedH = String(Math.floor(revisedMin / 60)).padStart(2, '0');
    const revisedM = String(revisedMin % 60).padStart(2, '0');
    const visitTimeRevised = `${revisedH}:${revisedM}`;

    // 알림톡 발송 — field_delayed
    const phoneNorm = (c.phone || '').replace(/\D/g, '');
    await sendNotification({
      template: 'field_delayed',
      phone: phoneNorm,
      name: c.name,
      data: {
        id: c.unique_id || c.id,
        name: c.name,
        phone: phoneNorm,
        type: '출장 요청',
        date: c.visit_date || '',
        time: c.visit_time || '',
        delay_min: String(delayMin),
        visit_time_revised: visitTimeRevised,
      },
    });

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
