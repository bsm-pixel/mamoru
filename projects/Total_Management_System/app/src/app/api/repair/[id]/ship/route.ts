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

    // ALPS 직접 호출 — 송장번호 발급 + 접수
    const { invoiceNumber } = await getNextInvoice();
    const fullAddress = [repair.address1, repair.address2].filter(Boolean).join(' ');

    const result = await bookShipment({
      invoiceNumber,
      receiverName: repair.name,
      receiverTel: repair.phone || '',
      receiverZip: repair.postcode || '',
      receiverAddr: fullAddress,
      goodsName: '가위 복원수리',
    });

    if (!result.success) {
      return NextResponse.json(
        { error: `ALPS 송장 생성 실패: ${result.error}` },
        { status: 502 }
      );
    }

    // Supabase 업데이트
    const { data: updated, error: updateErr } = await db
      .from('repairs')
      .update({
        status: 'ready_to_ship',
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
      to_status: 'ready_to_ship',
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

    // DB 상태 되돌림
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
