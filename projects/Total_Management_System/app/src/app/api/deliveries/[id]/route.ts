import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateImwebStock } from '@/lib/imweb/client';

/** GET /api/deliveries/[id] — 납품 상세 */
export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const [dlRes, itemsRes] = await Promise.all([
      db.from('deliveries').select('*').eq('id', id).single(),
      db.from('delivery_items').select('*').eq('delivery_id', id),
    ]);

    if (dlRes.error) throw dlRes.error;

    return NextResponse.json({
      delivery: dlRes.data,
      items: itemsRes.data || [],
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/deliveries/[id] — 상태 전환 + 편집 */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();
    const { action } = body as { action: string };

    // 납품 건 조회
    const { data: dl, error: dlErr } = await db.from('deliveries').select('*').eq('id', id).single();
    if (dlErr || !dl) return NextResponse.json({ error: '납품 건을 찾을 수 없습니다' }, { status: 404 });

    // ── 납품 확정 (draft → confirmed): 재고 차감 ──
    if (action === 'confirm') {
      if (dl.status !== 'draft') return NextResponse.json({ error: '작성중 상태에서만 확정 가능합니다' }, { status: 400 });

      // 재고 차감
      const { data: items } = await db.from('delivery_items').select('product_id, quantity').eq('delivery_id', id);
      if (items) {
        const qtyMap: Record<string, number> = {};
        for (const item of items) {
          if (item.product_id) qtyMap[item.product_id] = (qtyMap[item.product_id] || 0) + item.quantity;
        }
        for (const [productId, qty] of Object.entries(qtyMap)) {
          const { data: prod } = await db.from('products').select('stock_quantity, raw_stock, imweb_product_no').eq('id', productId).single();
          if (!prod) continue;
          const newStock = Math.max(0, (prod.stock_quantity || 0) - qty);
          const newRaw = Math.max(0, (prod.raw_stock || 0) - qty);
          await db.from('products').update({ stock_quantity: newStock, raw_stock: newRaw }).eq('id', productId);

          // 아임웹 동기화
          if (prod.imweb_product_no) {
            try {
              const { count: displayCount } = await db.from('product_serials')
                .select('id', { count: 'exact', head: true })
                .eq('product_id', productId).eq('status', 'in_stock').eq('warehouse_zone', 'display');
              await updateImwebStock(Number(prod.imweb_product_no), Math.max(0, newStock - (displayCount || 0)));
            } catch (e) { console.error('[imweb] 납품 재고 동기화 실패:', e); }
          }
        }
      }

      await db.from('deliveries').update({ status: 'confirmed', updated_at: new Date().toISOString() }).eq('id', id);
      return NextResponse.json({ success: true, status: 'confirmed' });
    }

    // ── 출고 완료 (confirmed → shipped) ──
    if (action === 'ship') {
      if (dl.status !== 'confirmed') return NextResponse.json({ error: '납품확정 상태에서만 출고 가능합니다' }, { status: 400 });
      const trackingNumber = body.tracking_number || null;
      await db.from('deliveries').update({
        status: 'shipped',
        shipped_date: new Date().toISOString().slice(0, 10),
        tracking_number: trackingNumber,
        updated_at: new Date().toISOString(),
      }).eq('id', id);
      return NextResponse.json({ success: true, status: 'shipped' });
    }

    // ── 정산 완료 (shipped → settled) ──
    if (action === 'settle') {
      if (dl.status !== 'shipped') return NextResponse.json({ error: '출고완료 상태에서만 정산 가능합니다' }, { status: 400 });

      // 미수금 차감 (잔액이 있으면)
      if (dl.customer_id && dl.payment_status !== 'paid') {
        const unpaid = dl.total_amount - (dl.discount_amount || 0) - (dl.paid_amount || 0);
        if (unpaid > 0) {
          const { data: cust } = await db.from('customers').select('outstanding_balance').eq('id', dl.customer_id).single();
          if (cust) {
            await db.from('customers').update({
              outstanding_balance: Math.max(0, (cust.outstanding_balance || 0) - unpaid),
            }).eq('id', dl.customer_id);
          }
        }
      }

      await db.from('deliveries').update({
        status: 'settled',
        payment_status: 'paid',
        paid_amount: dl.total_amount - (dl.discount_amount || 0),
        updated_at: new Date().toISOString(),
      }).eq('id', id);
      return NextResponse.json({ success: true, status: 'settled' });
    }

    // ── 취소 (draft/confirmed → cancelled) ──
    if (action === 'cancel') {
      if (dl.status === 'shipped' || dl.status === 'settled') {
        return NextResponse.json({ error: '출고/정산 완료 건은 취소할 수 없습니다' }, { status: 400 });
      }
      if (dl.cancelled_at) return NextResponse.json({ error: '이미 취소된 납품입니다' }, { status: 400 });

      const reason = body.reason || '';

      // confirmed였다면 재고 복원
      if (dl.status === 'confirmed') {
        const { data: items } = await db.from('delivery_items').select('product_id, quantity').eq('delivery_id', id);
        if (items) {
          for (const item of items) {
            if (!item.product_id) continue;
            const { data: prod } = await db.from('products').select('stock_quantity, raw_stock, imweb_product_no').eq('id', item.product_id).single();
            if (!prod) continue;
            const newStock = (prod.stock_quantity || 0) + item.quantity;
            const newRaw = (prod.raw_stock || 0) + item.quantity;
            await db.from('products').update({ stock_quantity: newStock, raw_stock: newRaw }).eq('id', item.product_id);

            if (prod.imweb_product_no) {
              try {
                const { count: displayCount } = await db.from('product_serials')
                  .select('id', { count: 'exact', head: true })
                  .eq('product_id', item.product_id).eq('status', 'in_stock').eq('warehouse_zone', 'display');
                await updateImwebStock(Number(prod.imweb_product_no), Math.max(0, newStock - (displayCount || 0)));
              } catch (e) { console.error('[imweb] 납품취소 재고 동기화 실패:', e); }
            }
          }
        }
      }

      // 미수금 차감
      if (dl.customer_id && dl.payment_status !== 'paid') {
        const unpaid = dl.total_amount - (dl.discount_amount || 0) - (dl.paid_amount || 0);
        if (unpaid > 0) {
          const { data: cust } = await db.from('customers').select('outstanding_balance').eq('id', dl.customer_id).single();
          if (cust) {
            await db.from('customers').update({
              outstanding_balance: Math.max(0, (cust.outstanding_balance || 0) - unpaid),
            }).eq('id', dl.customer_id);
          }
        }
      }

      await db.from('deliveries').update({
        cancelled_at: new Date().toISOString(),
        cancelled_reason: reason,
        updated_at: new Date().toISOString(),
      }).eq('id', id);
      return NextResponse.json({ success: true, action: 'cancelled' });
    }

    // ── 결제상태 변경 ──
    if (action === 'update_payment') {
      const { payment_status: newStatus, paid_amount: newPaid } = body as {
        payment_status: string; paid_amount?: number;
      };
      const oldUnpaid = dl.total_amount - (dl.discount_amount || 0) - (dl.paid_amount || 0);
      const newPaidVal = newStatus === 'paid' ? (dl.total_amount - (dl.discount_amount || 0)) : (newPaid || 0);
      const newUnpaid = dl.total_amount - (dl.discount_amount || 0) - newPaidVal;
      const diff = oldUnpaid - newUnpaid;

      await db.from('deliveries').update({
        payment_status: newStatus,
        paid_amount: newPaidVal,
        updated_at: new Date().toISOString(),
      }).eq('id', id);

      // 미수금 조정
      if (dl.customer_id && diff !== 0) {
        const { data: cust } = await db.from('customers').select('outstanding_balance').eq('id', dl.customer_id).single();
        if (cust) {
          await db.from('customers').update({
            outstanding_balance: Math.max(0, (cust.outstanding_balance || 0) - diff),
          }).eq('id', dl.customer_id);
        }
      }

      return NextResponse.json({ success: true });
    }

    // ── 메모/기타 필드 수정 ──
    const updates: Record<string, unknown> = {};
    if (body.memo !== undefined) updates.memo = body.memo;
    if (body.expected_date !== undefined) updates.expected_date = body.expected_date;
    if (body.tracking_number !== undefined) updates.tracking_number = body.tracking_number;
    if (Object.keys(updates).length > 0) {
      updates.updated_at = new Date().toISOString();
      await db.from('deliveries').update(updates).eq('id', id);
    }

    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
