import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

/**
 * POST /api/repair/[id]/merged-ship
 *
 * 복원수리 합포장 출고 처리:
 *   - 같은 고객의 판매건/주문건 송장에 복원수리 합쳐 발송한 케이스
 *   - 선택한 송장번호 + courier 를 repairs 에 복사
 *   - status='shipped' 전환 + shipped_at 설정
 *   - 기존 as_shipped 알림톡 자동 발송 (변경 없이 invoice_number 채워졌으므로 정상)
 *
 * Body:
 *   {
 *     invoice_number: string,    // 합산 송장번호 (필수)
 *     courier_name?: string,     // 택배사 (기본: '롯데택배')
 *     source_type: 'sale' | 'order' | 'manual',  // 출처 타입
 *     source_id?: string,        // offline_sales.id 또는 orders.id (manual 시 null)
 *     skip_notify?: boolean,     // true 면 알림톡 미발송 (옵션, 기본 false)
 *   }
 */
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

    const body = await req.json();
    const { invoice_number, courier_name, source_type, source_id, skip_notify } = body;

    // 입력 검증
    if (!invoice_number || !String(invoice_number).trim()) {
      return NextResponse.json({ error: 'invoice_number_required' }, { status: 400 });
    }
    if (!['sale', 'order', 'manual'].includes(source_type)) {
      return NextResponse.json({ error: 'invalid_source_type' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 현재 repair 상태 검증
    const { data: current, error: fetchErr } = await db
      .from('repairs')
      .select('*')
      .eq('id', id)
      .single();

    if (fetchErr || !current) {
      return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
    }

    if (current.status !== 'ready_to_ship') {
      return NextResponse.json(
        { error: 'invalid_status_transition', current: current.status, expected: 'ready_to_ship' },
        { status: 400 }
      );
    }

    const now = new Date().toISOString();
    const cleanInvoice = String(invoice_number).trim();
    const cleanCourier = (courier_name && String(courier_name).trim()) || '롯데택배';

    // 합포장 출처 정보 박제용 note (자동 추적 가능)
    const sourceLabel = source_type === 'sale'
      ? `오프라인판매(${source_id || '-'})`
      : source_type === 'order'
      ? `아임웹주문(${source_id || '-'})`
      : '직접입력';
    const note = `판매건 합포장 출고 — 송장 ${cleanInvoice} (${cleanCourier}, 출처: ${sourceLabel})`;

    // 업데이트: 송장 복사 + 상태 전환
    const { data: updated, error: updateErr } = await db
      .from('repairs')
      .update({
        status: 'shipped',
        invoice_number: cleanInvoice,
        courier_name: cleanCourier,
        shipped_at: now,
      })
      .eq('id', id)
      .select()
      .single();

    if (updateErr) throw updateErr;

    // 이력 기록
    await db.from('repair_history').insert({
      repair_id: id,
      from_status: current.status,
      to_status: 'shipped',
      changed_by: user.id,
      note,
    });

    // 알림톡 발송 (기본 발송, skip_notify=true 시 우회)
    if (!skip_notify) {
      after(async () => {
        if (!updated.phone) return;
        const result = await sendNotification({
          template: 'as_shipped',
          phone: updated.phone,
          name: updated.name,
          data: {
            id: updated.as_id,
            as_amount: String(updated.service_cost || 0),
            shipping_amount: String(updated.shipping_fee || 0),
            total_amount: String(updated.total_amount || 0),
            tracking: cleanInvoice,
            courier: cleanCourier,
            uid: updated.as_id,
            as_uid: updated.as_id,
            review_type: 'repair',
            type_label: '복원수리',
          },
        });
        if (!result.success) {
          console.error('[merged-ship notify] 실패:', result.error);
        } else {
          console.log('[merged-ship notify] as_shipped 발송 성공');
        }
      });
    }

    return NextResponse.json({ ok: true, repair: updated, note });
  } catch (err) {
    console.error('[merged-ship] 처리 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
