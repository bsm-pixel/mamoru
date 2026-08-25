import { NextRequest, NextResponse } from 'next/server';
import { after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { isValidReturnTransition } from '@/lib/returns/transitions';
import { sendNotification } from '@/lib/notification/make-webhook';
import type { ReturnStatus } from '@/lib/supabase/types';

/** 상태 → 채울 타임스탬프 컬럼 */
const STATUS_TS: Partial<Record<ReturnStatus, string>> = {
  pickup_scheduled: 'pickup_scheduled_at',
  inbound: 'inbound_at',
  inspected: 'inspected_at',
  completed: 'completed_at',
  cancelled: 'cancelled_at',
};

/** PATCH /api/returns/[id] — 상태 전이 / 필드 수정 */
export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();

    const { data: cur } = await db.from('returns').select('*').eq('id', id).single();
    if (!cur) return NextResponse.json({ error: '반품 건을 찾을 수 없습니다' }, { status: 404 });

    const update: Record<string, unknown> = {};

    // 상태 전이
    if (body.status && body.status !== cur.status) {
      if (!isValidReturnTransition(cur.status, body.status)) {
        return NextResponse.json({ error: `전이 불가: ${cur.status} → ${body.status}` }, { status: 400 });
      }
      update.status = body.status;
      const tsCol = STATUS_TS[body.status as ReturnStatus];
      if (tsCol && !cur[tsCol]) update[tsCol] = new Date().toISOString();
      if (body.status === 'cancelled' && body.cancelled_reason) update.cancelled_reason = body.cancelled_reason;

      // 입고완료 시: 구 시리얼을 반품창고(returned/return)로 확정 (idempotent — 교환 시 이미 이동됐으면 무해)
      if (body.status === 'inbound' && cur.serial_id) {
        await db.from('product_serials').update({
          status: 'returned', warehouse_zone: 'return',
          offline_sale_id: null, sale_item_id: null, sold_via: null, sold_at: null,
          sold_to_name: null, sold_to_phone: null,
        }).eq('id', cur.serial_id);
      }
    }

    // 수정 가능 필드 (수거 예약일·방식·송장·메모 등)
    for (const f of ['pickup_method', 'pickup_date', 'courier_name', 'invoice_number', 'memo', 'admin_note', 'refund_amount', 'refund_method']) {
      if (f in body) update[f] = body[f];
    }

    if (Object.keys(update).length === 0) return NextResponse.json({ ok: true });

    const { data, error } = await db.from('returns').update(update).eq('id', id).select().single();
    if (error) throw error;

    // 입고완료 → 고객에게 "반품 잘 받았습니다" 알림톡(솔라피 return_inbound 등록 시). after()로 완주 보장
    if (update.status === 'inbound' && cur.phone) {
      after(async () => {
        try {
          await sendNotification({
            template: 'return_inbound',
            phone: String(cur.phone),
            name: String(cur.name || ''),
            data: { return_number: String(cur.return_number || ''), product_name: String(cur.product_name || '') },
          });
        } catch (e) { console.error('[returns PATCH] 입고 알림 실패:', e); }
      });
    }

    return NextResponse.json({ return: data });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
