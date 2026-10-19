import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getNextInvoice, bookShipment, cancelShipment } from '@/lib/lotte/alps-client';

/** POST /api/repair/[id]/ship — 송장 생성 (ALPS 직접 호출) */
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

    if (!repair.phone) {
      return NextResponse.json({ error: '고객 연락처가 없습니다. 수리 접수 정보를 확인해주세요.' }, { status: 400 });
    }

    // ALPS 직접 호출 — 송장번호 발급 + 접수
    const { invoiceNumber } = await getNextInvoice();
    const fullAddress = [repair.address, repair.address_detail].filter(Boolean).join(' ');

    const result = await bookShipment({
      invoiceNumber,
      receiverName: repair.name,
      receiverTel: repair.phone || '',
      receiverZip: repair.postcode || '',
      receiverAddr: fullAddress,
      goodsName: '[MAMORU] 복원수리',
    });

    if (!result.success) {
      return NextResponse.json(
        { error: `ALPS 송장 생성 실패: ${result.error}` },
        { status: 502 }
      );
    }

    // 2026-10-19: 송장은 작업 상태와 별개 — 미리 뽑아도 진행중이면 진행중 그대로 (작업완료는 사장님이 따로)
    //   기사님 수거는 상태와 무관하게 크론 track-delivery [2-A] 가 감지해 출고완료 + 알림톡
    const { data: updated, error: updateErr } = await db
      .from('repairs')
      .update({
        invoice_number: invoiceNumber,
        courier_name: '롯데택배',
      })
      .eq('id', id)
      .select()
      .single();

    if (updateErr) throw updateErr;

    await db.from('repair_history').insert({
      repair_id: id,
      from_status: repair.status,
      to_status: repair.status,
      changed_by: user.id,
      note: `송장 생성: ${invoiceNumber}`,
    });

    return NextResponse.json({ repair: updated, invNo: invoiceNumber });
  } catch (err) {
    console.error('[repair] 출고 실패:', err);
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}

/** DELETE /api/repair/[id]/ship — 송장 취소 (ALPS + DB) */
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

    // ALPS 취소 시도
    let alpsWarning: string | undefined;
    if (cancelledInvNo) {
      const cancelResult = await cancelShipment(cancelledInvNo);
      if (!cancelResult.success) {
        alpsWarning = `ALPS 취소 실패: ${cancelResult.error} — 수동 취소 필요`;
      }
    }

    // 2026-10-19: 송장만 지운다 — 작업 상태는 그대로 (전엔 무조건 진행중으로 되돌려 작업완료 표시가 날아갔다)
    //   출고완료(shipped)에서 취소하면 아직 안 나간 것이므로 작업완료로
    const nextStatus = repair.status === 'shipped' ? 'ready_to_ship' : repair.status;
    const { data: updated } = await db
      .from('repairs')
      .update({
        status: nextStatus,
        invoice_number: null,
        courier_name: null,
        shipped_at: null,
        shipped_source: null,
      })
      .eq('id', id)
      .select()
      .single();

    await db.from('repair_history').insert({
      repair_id: id,
      from_status: repair.status,
      to_status: nextStatus,
      changed_by: user.id,
      note: `송장 취소: ${cancelledInvNo}${alpsWarning ? ' (' + alpsWarning + ')' : ''}`,
    });

    return NextResponse.json({
      ...updated,
      warning: alpsWarning,
    });
  } catch (err) {
    console.error('[repair] 송장 취소 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
