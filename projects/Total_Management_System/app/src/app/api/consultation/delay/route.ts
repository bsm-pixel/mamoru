import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/consultation/delay — 출장 지연 안내 (GAS fieldDelay 액션 호출) */
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

    // GAS fieldDelay 액션 호출
    const baseUrl = process.env.GAS_CONSULTING_URL;
    if (!baseUrl) {
      return NextResponse.json({ error: 'GAS_CONSULTING_URL 미설정' }, { status: 500 });
    }

    const key = process.env.CRON_SECRET || 'mamoru-tms-cron-2026';
    const params = new URLSearchParams({
      action: 'fieldDelay',
      uid: c.unique_id,
      delayMin: String(delayMin),
      key,
    });

    const gasRes = await fetch(`${baseUrl}?${params.toString()}`, {
      method: 'GET',
      redirect: 'follow',
      signal: AbortSignal.timeout(15000),
    });

    const text = await gasRes.text();
    let gasBody;
    try { gasBody = JSON.parse(text); } catch { gasBody = text; }

    if (!gasBody?.ok) {
      console.error('[delay] GAS 실패:', gasBody);
      return NextResponse.json({ error: gasBody?.error || 'GAS 호출 실패' }, { status: 502 });
    }

    // 이력 기록
    await db.from('consultation_history').insert({
      consultation_id: consultationId,
      from_status: c.status,
      to_status: c.status, // 상태 변경 없이 이력만 기록
      changed_by: user.id,
      note: `출장 지연 안내: ${delayMin}분 (도착 예정 ${gasBody.visit_time_revised})`,
    });

    return NextResponse.json({
      success: true,
      delay_min: delayMin,
      visit_time_revised: gasBody.visit_time_revised,
    });
  } catch (err) {
    console.error('[delay] 출장 지연 안내 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
