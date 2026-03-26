import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/repair/[id]/ship — 출고 처리 (GAS ALPS 경유) */
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

    // 현재 복원수리 조회
    const { data: repair, error: fetchErr } = await db
      .from('repairs')
      .select('*')
      .eq('id', id)
      .single();

    if (fetchErr || !repair) {
      return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
    }

    if (!['cost_notified', 'repairing', 'ready_to_ship'].includes(repair.status)) {
      return NextResponse.json({ error: `출고 불가 상태: ${repair.status}` }, { status: 400 });
    }

    // GAS ALPS 호출
    const gasUrl = process.env.GAS_AS_URL;
    const adminToken = process.env.GAS_AS_ADMIN_TOKEN;
    if (!gasUrl || !adminToken) {
      return NextResponse.json({ error: 'GAS_AS_URL 또는 GAS_AS_ADMIN_TOKEN 미설정' }, { status: 500 });
    }

    const gasParams = new URLSearchParams({
      action: 'book',
      as_id: repair.as_id,
      token: adminToken,
    });

    const gasRes = await fetch(`${gasUrl}?${gasParams}`, {
      method: 'GET',
      redirect: 'follow',
      signal: AbortSignal.timeout(30000),
    });

    const gasText = await gasRes.text();
    let gasBody;
    try { gasBody = JSON.parse(gasText); } catch { gasBody = { ok: false, msg: gasText }; }

    if (!gasBody.ok && !gasBody.invNo) {
      return NextResponse.json(
        { error: `GAS 송장 생성 실패: ${gasBody.msg || gasText}` },
        { status: 502 }
      );
    }

    // Supabase 업데이트 — 송장 생성 = ready_to_ship (출고완료는 별도 액션)
    const { data: updated, error: updateErr } = await db
      .from('repairs')
      .update({
        status: 'ready_to_ship',
        invoice_number: gasBody.invNo,
        courier_name: '롯데택배',
      })
      .eq('id', id)
      .select()
      .single();

    if (updateErr) throw updateErr;

    // 이력 기록
    await db.from('repair_history').insert({
      repair_id: id,
      from_status: repair.status,
      to_status: 'ready_to_ship',
      changed_by: user.id,
      note: `송장 생성: ${gasBody.invNo}`,
    });

    return NextResponse.json({
      repair: updated,
      invNo: gasBody.invNo,
    });
  } catch (err) {
    console.error('[repair] 출고 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/repair/[id]/ship — 송장 취소 (DB만 처리, ALPS는 수동 취소) */
export async function DELETE(
  _req: NextRequest,
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

    const { data: repair } = await db
      .from('repairs')
      .select('as_id, invoice_number, status')
      .eq('id', id)
      .single();

    if (!repair) {
      return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
    }

    const cancelledInvNo = repair.invoice_number || '';

    // DB 상태 되돌림 (ALPS 취소 API 미지원 → 수동 취소 필요)
    const { data: updated } = await db
      .from('repairs')
      .update({
        status: 'repairing',
        invoice_number: null,
        shipped_at: null,
      })
      .eq('id', id)
      .select()
      .single();

    await db.from('repair_history').insert({
      repair_id: id,
      from_status: repair.status,
      to_status: 'repairing',
      changed_by: user.id,
      note: `송장 취소: ${cancelledInvNo} (ALPS 수동 취소 필요)`,
    });

    return NextResponse.json({
      ...updated,
      warning: cancelledInvNo
        ? `송장 ${cancelledInvNo}은 ALPS에서 직접 취소해주세요.`
        : undefined,
    });
  } catch (err) {
    console.error('[repair] 송장 취소 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
