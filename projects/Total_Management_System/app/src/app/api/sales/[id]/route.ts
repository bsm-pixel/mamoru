import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateImwebStock } from '@/lib/imweb/client';

/** PATCH /api/sales/[id] — 취소 / 결제상태 변경 / 메모 수정 */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
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

    // 현재 판매 건 조회
    const { data: sale, error: saleErr } = await db
      .from('offline_sales')
      .select('*')
      .eq('id', id)
      .single();
    if (saleErr || !sale) {
      return NextResponse.json({ error: '판매 건을 찾을 수 없습니다' }, { status: 404 });
    }

    // --- A) 판매 취소 ---
    if (action === 'cancel') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '이미 취소된 판매입니다' }, { status: 400 });
      }

      const reason = body.reason || '';

      // 1. 시리얼 복원 — previous_zone으로 원래 위치 복원
      const { data: serials } = await db
        .from('product_serials')
        .select('id, product_id, previous_zone')
        .eq('offline_sale_id', id);

      if (serials && serials.length > 0) {
        for (const serial of serials) {
          await db
            .from('product_serials')
            .update({
              status: 'in_stock',
              warehouse_zone: serial.previous_zone || 'ready', // 원래 zone 복원 (NULL이면 ready fallback)
              previous_zone: null,
              sold_via: null,
              offline_sale_id: null,
              sold_at: null,
              sold_to_name: null,
              sold_to_phone: null,
            })
            .eq('id', serial.id);
        }
      }

      // 2. 재고 복원 + 3. 아임웹 재고 동기화
      const { data: items } = await db
        .from('offline_sale_items')
        .select('product_id, quantity')
        .eq('sale_id', id);

      if (items && items.length > 0) {
        // 상품별 수량 집계 후 병렬 처리
        const productQtyMap: Record<string, number> = {};
        for (const item of items) {
          if (item.product_id && item.quantity) {
            productQtyMap[item.product_id] = (productQtyMap[item.product_id] || 0) + item.quantity;
          }
        }

        await Promise.all(Object.entries(productQtyMap).map(async ([productId, qty]) => {
          const { data: prod } = await db
            .from('products')
            .select('stock_quantity, imweb_product_no')
            .eq('id', productId)
            .single();
          if (!prod) return;

          const newStock = (prod.stock_quantity || 0) + qty;
          await db.from('products').update({ stock_quantity: newStock }).eq('id', productId);

          if (prod.imweb_product_no) {
            try {
              // 디스플레이 제외 판매 가능 재고만 아임웹에 전달
              const { count: displayCount } = await db
                .from('product_serials')
                .select('id', { count: 'exact', head: true })
                .eq('product_id', productId)
                .eq('status', 'in_stock')
                .eq('warehouse_zone', 'display');
              const sellableStock = Math.max(0, newStock - (displayCount || 0));
              await updateImwebStock(Number(prod.imweb_product_no), sellableStock);
            } catch (e) {
              console.error('[imweb] 취소 재고 동기화 실패:', prod.imweb_product_no, e);
            }
          }
        }));
      }

      // 4. 미수금 차감 (미결제/부분결제였던 경우)
      if (sale.customer_id && sale.payment_status !== 'paid') {
        const unpaidAmount = sale.total_amount - (sale.paid_amount || 0);
        if (unpaidAmount > 0) {
          const { data: cust } = await db
            .from('customers')
            .select('outstanding_balance')
            .eq('id', sale.customer_id)
            .single();
          if (cust) {
            await db
              .from('customers')
              .update({
                outstanding_balance: Math.max(0, (cust.outstanding_balance || 0) - unpaidAmount),
              })
              .eq('id', sale.customer_id);
          }
        }
      }

      // 5. 판매 상태 업데이트
      const { error: updateErr } = await db
        .from('offline_sales')
        .update({
          cancelled_at: new Date().toISOString(),
          cancelled_reason: reason,
          cancelled_by: user.id,
        })
        .eq('id', id);

      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'cancelled' });
    }

    // --- B) 결제상태 변경 ---
    if (action === 'update_payment') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매는 변경할 수 없습니다' }, { status: 400 });
      }

      const { payment_status, paid_amount } = body as {
        payment_status: string;
        paid_amount?: number;
      };

      if (!payment_status) {
        return NextResponse.json({ error: 'payment_status 필수' }, { status: 400 });
      }

      // 미수금 조정
      if (sale.customer_id) {
        const oldUnpaid = sale.payment_status !== 'paid'
          ? sale.total_amount - (sale.paid_amount || 0) : 0;
        const newPaidAmount = paid_amount ?? (payment_status === 'paid' ? sale.total_amount : sale.paid_amount);
        const newUnpaid = payment_status !== 'paid'
          ? sale.total_amount - newPaidAmount : 0;
        const diff = oldUnpaid - newUnpaid; // 양수면 미수금 감소

        if (diff !== 0) {
          const { data: cust } = await db
            .from('customers')
            .select('outstanding_balance')
            .eq('id', sale.customer_id)
            .single();
          if (cust) {
            await db
              .from('customers')
              .update({
                outstanding_balance: Math.max(0, (cust.outstanding_balance || 0) - diff),
              })
              .eq('id', sale.customer_id);
          }
        }
      }

      const updateData: Record<string, unknown> = { payment_status };
      if (paid_amount !== undefined) updateData.paid_amount = paid_amount;

      const { error: updateErr } = await db
        .from('offline_sales')
        .update(updateData)
        .eq('id', id);

      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'payment_updated' });
    }

    // --- C) 메모 수정 ---
    if (action === 'update_memo') {
      const { memo } = body as { memo: string };
      const { error: updateErr } = await db
        .from('offline_sales')
        .update({ memo: memo || null })
        .eq('id', id);

      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'memo_updated' });
    }

    return NextResponse.json({ error: '알 수 없는 action' }, { status: 400 });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    console.error('[sales PATCH] error:', message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
