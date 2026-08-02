import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';
import { convertEventToSale } from '@/lib/event/convert-to-sale';

/** GET /api/events/[id] */
export async function GET(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (db as any).from('event_submissions').select('*').eq('id', id).single();
    if (error) throw error;
    return NextResponse.json({ ok: true, event: data });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/events/[id] — action: payment_notice | confirm_payment | cancel | update */
export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const body = await req.json();
    const action = body.action as string;
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const { data: ev, error: getErr } = await dbAny.from('event_submissions').select('*').eq('id', id).single();
    if (getErr || !ev) throw getErr || new Error('접수 없음');

    const phoneNorm = (ev.customer_phone || '').replace(/\D/g, '');

    // 117: 재고판매(stock_sale)는 EVENT 와 알림톡 템플릿·판매전환 라벨을 달리 쓴다
    const isStock = ev.kind === 'stock_sale';
    const T = isStock
      ? { notice: 'stock_payment_notice' as const, confirmed: 'stock_payment_confirmed' as const }
      : { notice: 'event_payment_notice' as const, confirmed: 'event_payment_confirmed' as const };
    const convertOpts = isStock ? { category: 'LS', memoLabel: '재고판매 전환' } : undefined;
    const numKey = isStock ? 'order_number' : 'event_number';

    if (action === 'payment_notice') {
      // 사장님 재고확인 후 입금안내 (총액 확정 — 금액 조정 가능)
      const total = typeof body.total_amount === 'number' ? body.total_amount : ev.total_amount;
      await dbAny.from('event_submissions').update({
        status: 'payment_noticed', payment_noticed_at: new Date().toISOString(), total_amount: total,
      }).eq('id', id);
      await dbAny.from('event_history').insert({ event_id: id, to_status: 'payment_noticed', note: '입금안내 발송' });
      if (phoneNorm && body.send_notification !== false) {
        after(async () => {
          await sendNotification({
            template: T.notice, phone: phoneNorm, name: ev.customer_name,
            data: { id: ev.event_number, [numKey]: ev.event_number, total_amount: String(total) },
          });
        });
      }
      return NextResponse.json({ ok: true });
    }

    if (action === 'confirm_payment') {
      // 입금확인 → 판매 자동 전환 (EVENT/재고판매 공통 파이프라인, 라벨만 분기)
      const saleId = await convertEventToSale(dbAny, ev, convertOpts);
      await dbAny.from('event_submissions').update({
        status: 'converted', paid_at: new Date().toISOString(), sale_id: saleId,
      }).eq('id', id);
      await dbAny.from('event_history').insert({ event_id: id, to_status: 'converted', note: `입금확인 → 판매 전환` });
      if (phoneNorm && body.send_notification !== false) {
        after(async () => {
          await sendNotification({
            template: T.confirmed, phone: phoneNorm, name: ev.customer_name,
            data: { id: ev.event_number, [numKey]: ev.event_number, total_amount: String(ev.total_amount) },
          });
        });
      }
      return NextResponse.json({ ok: true, sale_id: saleId });
    }

    if (action === 'cancel') {
      // 🔴 이미 판매전환된 건은, 연결 판매가 아직 '살아있으면' 접수만 취소하면 판매가 유령으로 남아
      //    회계(매출)에 계속 잡힌다. → 판매관리에서 판매를 먼저 취소/반품해야 재고·환불·회계가 함께 정리됨.
      //    (판매가 이미 취소/반품됐거나 sale_id 없으면 접수 기록 정리는 그대로 허용)
      if (ev.sale_id) {
        const { data: sale } = await dbAny
          .from('offline_sales')
          .select('cancelled_at, returned_at, sale_number')
          .eq('id', ev.sale_id)
          .single();
        if (sale && !sale.cancelled_at && !sale.returned_at) {
          return NextResponse.json({
            ok: false,
            code: 'active_sale',
            sale_id: ev.sale_id,
            error: `이미 판매(${sale.sale_number || ''})로 전환된 접수입니다.\n판매관리에서 판매를 먼저 취소/반품해 주세요 — 그래야 재고·환불·회계가 함께 정리됩니다.`,
          }, { status: 409 });
        }
      }
      await dbAny.from('event_submissions').update({
        status: 'cancelled', cancelled_at: new Date().toISOString(),
      }).eq('id', id);
      await dbAny.from('event_history').insert({ event_id: id, to_status: 'cancelled', note: body.reason || '취소' });
      return NextResponse.json({ ok: true });
    }

    if (action === 'update') {
      const patch: Record<string, unknown> = {};
      if (typeof body.memo === 'string') patch.memo = body.memo;
      if (typeof body.total_amount === 'number') patch.total_amount = body.total_amount;
      if (Object.keys(patch).length > 0) await dbAny.from('event_submissions').update(patch).eq('id', id);
      return NextResponse.json({ ok: true });
    }

    return NextResponse.json({ ok: false, error: '알 수 없는 action' }, { status: 400 });
  } catch (err) {
    console.error('[events PATCH] 실패:', err);
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}

/**
 * DELETE /api/events/[id] — 접수 기록 완전 삭제 (event_history 포함).
 *  ⚠️ 연결된 판매(sale_id)는 건드리지 않는다 — 판매는 판매관리에서 별도 취소/관리(재고 정합성은 그쪽 소관).
 *     오등록·테스트 접수 정리용.
 */
export async function DELETE(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;
    await dbAny.from('event_history').delete().eq('event_id', id);
    const { error } = await dbAny.from('event_submissions').delete().eq('id', id);
    if (error) throw error;
    return NextResponse.json({ ok: true });
  } catch (err) {
    console.error('[events DELETE] 실패:', err);
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
