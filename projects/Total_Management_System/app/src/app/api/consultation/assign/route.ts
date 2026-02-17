import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/consultation/assign — 딜러 배정 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { consultationId, dealerId } = await req.json();

    if (!consultationId || !dealerId) {
      return NextResponse.json(
        { error: 'consultationId와 dealerId가 필요합니다' },
        { status: 400 }
      );
    }

    // 현재 상태 조회
    const { data: current } = await db
      .from('consultations')
      .select('status')
      .eq('id', consultationId)
      .single();

    // 딜러 배정 + 상태 변경
    const { data, error } = await db
      .from('consultations')
      .update({
        dealer_id: dealerId,
        status: 'assigned',
      })
      .eq('id', consultationId)
      .select()
      .single();

    if (error) throw error;

    // 이력 기록
    await db.from('consultation_history').insert({
      consultation_id: consultationId,
      from_status: current?.status || null,
      to_status: 'assigned',
      changed_by: user.id,
      note: '딜러 배정',
    });

    return NextResponse.json(data);
  } catch (err) {
    console.error('[consultation] 딜러 배정 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
