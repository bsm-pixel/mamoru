import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { normalizeChecks } from '@/lib/repair/wound-catalog';

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

    // 전체 교체 방식 — 2026-10-15: "삭제 후 삽입"은 삽입 실패 시 기존 내역서가 통째로 사라졌다.
    //   → 기존 행 id를 먼저 잡고, 새 행 삽입이 성공한 뒤에만 기존 행을 지운다 (실패 시 기존 보존)
    const { data: oldRows, error: oldErr } = await db
      .from('repair_inspections').select('id').eq('repair_id', id);
    if (oldErr) throw oldErr;
    const oldIds: string[] = (oldRows || []).map((r: { id: string }) => r.id);

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
      // 163 v3: NULL = 구 사진핀 내역서 그대로 / 배열 = 그림체 체크 id (모르는 id 는 버림)
      diagram_marks: insp.diagram_marks == null ? null : normalizeChecks(insp.diagram_marks),
      zoom_photos: Array.isArray(insp.zoom_photos) ? (insp.zoom_photos as unknown[]).filter((u): u is string => typeof u === 'string').slice(0, 3) : [],
      worker: insp.worker || '백성민',
    }));

    const { data, error } = await db
      .from('repair_inspections')
      .insert(rows)
      .select();

    if (error) throw error;

    if (oldIds.length > 0) {
      const { error: delErr } = await db.from('repair_inspections').delete().in('id', oldIds);
      if (delErr) throw delErr;
    }

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
