import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * POST /api/repair/[id]/rework — 재수거품 입고 · 재작업 시작
 *   정밀 재점검으로 재수거(recall)한 출고건을 다시 작업 상태로 되돌린다.
 *   상태전이 검증(PATCH)이 막는 역방향(shipped/delivered/completed → repairing)을 전용 처리.
 *   흐름: [재수거 접수] → (회수) → 이 액션 → repairing → (재작업+내역서 갱신) → 송장생성 → 출고완료(as_shipped 재발화)
 *
 *   초기화: invoice_number/shipped_at/delivered_at/courier_name (재출고 준비).
 *   유지  : recall_invoice_number(재수거 이력)·paid_at·검수·비용.
 */
export async function POST(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data: r } = await db.from('repairs')
      .select('id, status, invoice_number, recall_invoice_number')
      .eq('id', id).single();
    if (!r) return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
    if (!r.recall_invoice_number) {
      return NextResponse.json({ error: '재수거 접수된 건만 재작업할 수 있습니다' }, { status: 400 });
    }
    if (!['shipped', 'delivered', 'completed'].includes(r.status)) {
      return NextResponse.json({ error: `출고 이후 상태에서만 재작업 가능(현재 ${r.status})` }, { status: 400 });
    }

    const now = new Date().toISOString();
    const { data, error } = await db.from('repairs').update({
      status: 'repairing',
      invoice_number: null,
      courier_name: null,
      shipped_at: null,
      delivered_at: null,
      reworked_at: now,   // 143: 재작업 시작 마커 → '재수리' 탭에서 빠지고 진행중 흐름으로 편입
      updated_at: now,
    }).eq('id', id).select().single();
    if (error) throw error;

    // 이력 기록 (이전 출고송장 보존)
    await db.from('repair_history').insert({
      repair_id: id,
      from_status: r.status,
      to_status: 'repairing',
      changed_by: user.id,
      note: `재수거품 입고 · 재작업 시작${r.invoice_number ? ` (이전 출고송장 ${r.invoice_number})` : ''}`,
    });

    return NextResponse.json({ ok: true, repair: data });
  } catch (err) {
    console.error('[repair/rework]', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
