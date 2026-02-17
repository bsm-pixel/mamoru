import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { isValidTransition } from '@/lib/consultation/transitions';
import { sendNotification } from '@/lib/notification/make-webhook';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';

/** POST /api/consultation/suggest — 시간 제안 (출장요청용) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const body = await req.json();
    const { consultationId, suggestions } = body as {
      consultationId: string;
      suggestions: { date: string; time: string }[];
    };

    if (!consultationId || !suggestions?.length) {
      return NextResponse.json({ error: 'consultationId, suggestions 필수' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data: current, error: fetchErr } = await db
      .from('consultations')
      .select('status, consultation_type, name, phone')
      .eq('id', consultationId)
      .single();

    if (fetchErr || !current) {
      return NextResponse.json({ error: '상담을 찾을 수 없습니다' }, { status: 404 });
    }

    // 상태 전이 검증: → suggested
    const canTransition = isValidTransition(
      current.consultation_type as ConsultationType,
      current.status as ConsultationStatus,
      'suggested'
    );
    if (!canTransition) {
      return NextResponse.json(
        { error: `현재 상태(${current.status})에서 시간 제안 불가` },
        { status: 400 }
      );
    }

    // suggestions JSONB 저장 + status → suggested
    const { data, error } = await db
      .from('consultations')
      .update({
        suggestions: { dates: suggestions },
        status: 'suggested',
      })
      .eq('id', consultationId)
      .select()
      .single();

    if (error) throw error;

    // 이력 기록
    await db.from('consultation_history').insert({
      consultation_id: consultationId,
      from_status: current.status,
      to_status: 'suggested',
      changed_by: user.id,
      note: `시간 제안: ${suggestions.map((s) => `${s.date} ${s.time}`).join(', ')}`,
    });

    // 알림톡 발송 (실패해도 API 성공 처리)
    const suggestText = suggestions.map((s) => `${s.date} ${s.time}`).join(' / ');
    await sendNotification({
      template: 'suggest',
      phone: current.phone,
      name: current.name,
      data: { suggestText },
    }).catch((err) => console.error('[suggest] 알림톡 실패:', err));

    return NextResponse.json(data);
  } catch (err) {
    console.error('[suggest] 시간 제안 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
