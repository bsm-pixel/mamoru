import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { cancelShipment } from '@/lib/lotte/alps-client';

/** DELETE /api/manual-invoices/[id] — 빠른 송장 취소 (soft delete + ALPS 취소 시도) */
export async function DELETE(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json().catch(() => ({}));
    const reason: string | null = (body.reason ?? '').trim() || null;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: invoice, error: fetchErr } = await db
      .from('manual_invoices')
      .select('id, invoice_number, cancelled_at')
      .eq('id', id)
      .single();

    if (fetchErr || !invoice) {
      return NextResponse.json({ error: '송장을 찾을 수 없습니다' }, { status: 404 });
    }
    if (invoice.cancelled_at) {
      return NextResponse.json({ error: '이미 취소된 송장입니다' }, { status: 400 });
    }

    // ALPS 취소 시도 (실패해도 DB soft delete는 진행)
    const cancelResult = await cancelShipment(invoice.invoice_number);
    const alpsFailed = !cancelResult.success;

    const { error: updateErr } = await db
      .from('manual_invoices')
      .update({
        cancelled_at: new Date().toISOString(),
        cancelled_by: user.id,
        cancelled_reason: reason,
        alps_cancel_failed: alpsFailed,
      })
      .eq('id', id);

    if (updateErr) {
      return NextResponse.json({ error: `DB 취소 처리 실패: ${updateErr.message}` }, { status: 500 });
    }

    return NextResponse.json({
      success: true,
      warning: alpsFailed ? `ALPS 취소 실패: ${cancelResult.error}. DB는 취소 처리되었지만 ALPS에서 직접 추가 확인이 필요할 수 있습니다.` : undefined,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
