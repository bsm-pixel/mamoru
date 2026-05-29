import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/repair/[id]/inspect — 검수 데이터 저장 (가위별) */
export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();

    // body: { inspections: [{ scissor_number, scissor_type, blade_tip, ... }] }
    const { inspections } = body;
    if (!inspections || !Array.isArray(inspections)) {
      return NextResponse.json({ error: 'inspections 배열 필수' }, { status: 400 });
    }

    // 기존 검수 데이터 삭제 후 재삽입 (전체 교체 방식)
    await db.from('repair_inspections').delete().eq('repair_id', id);

    const rows = inspections.map((insp: Record<string, unknown>) => ({
      repair_id: id,
      scissor_number: insp.scissor_number,
      scissor_type: insp.scissor_type || null,
      blade_tip: insp.blade_tip || '양호',
      blade_mid: insp.blade_mid || '양호',
      blade_inner: insp.blade_inner || '양호',
      comb: insp.comb || '',
      tension: insp.tension || '양호',
      parts: insp.parts || '양호',
      stopper: insp.stopper || '양호',
      photo_url: insp.photo_url || null,
      photo_marks: insp.photo_marks || null,
      comment: insp.comment ?? null,  // 097: 가위별 진단 및 내역
      worker: insp.worker || '백성민',
    }));

    const { data, error } = await db
      .from('repair_inspections')
      .insert(rows)
      .select();

    if (error) throw error;

    return NextResponse.json({ inspections: data });
  } catch (err) {
    console.error('[repair] 검수 저장 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PUT /api/repair/[id]/inspect — 검수 데이터 수정 (단건) */
export async function PUT(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();
    const { inspectionId, ...updateData } = body;

    if (!inspectionId) {
      return NextResponse.json({ error: 'inspectionId 필수' }, { status: 400 });
    }

    // repair_id 일치 확인
    const { data, error } = await db
      .from('repair_inspections')
      .update(updateData)
      .eq('id', inspectionId)
      .eq('repair_id', id)
      .select()
      .single();

    if (error) throw error;

    return NextResponse.json(data);
  } catch (err) {
    console.error('[repair] 검수 수정 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
